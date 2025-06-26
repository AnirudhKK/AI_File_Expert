import pandas as pd
import ollama
import yaml
import sys
import os
import re
import textwrap
import numpy as np
from datetime import datetime

# === Step 1: Get Excel File Path ===
file_path = input("🔍 Enter the path to your Excel file: ").strip()
if not os.path.exists(file_path):
    print("❌ File not found. Please check the path.")
    sys.exit(1)

# === Step 2: Get Template Name ===
template_name = input("📄 Enter the template name: ").strip()
template_file = re.sub(r'\s+', '_', template_name.strip().lower()) + ".txt"

# === Step 3: Load Excel File ===
try:
    df = pd.read_excel(file_path)
except Exception as e:
    print(f"❌ Failed to load Excel file: {e}")
    sys.exit(1)

# === If Template Exists: Preview + Options ===
if os.path.exists(template_file):
    print(f"\n📂 Found template file '{template_file}'.")

    with open(template_file, "r") as f:
        commands = [line.rstrip() for line in f]

    if not commands:
        print("⚠️ Template file is empty. Nothing to execute.")
        sys.exit(0)

    print("\n📖 Current template commands:\n")
    for idx, cmd in enumerate(commands, 1):
        print(f"{idx:>2}: {cmd}")

    # Prompt for edit/delete/explain
    while True:
        action = input("\n✏️ Do you want to modify this template? (y = edit / d = delete lines / e = explain / enter to continue): ").strip().lower()
        if action == 'y':
            print("📝 Editing full template content. End with ':wq' on a new line.\n")
            new_lines = []
            while True:
                line = input()
                if line.strip() == ":wq":
                    break
                new_lines.append(line)
            with open(template_file, "w") as f:
                f.write("\n".join(new_lines) + "\n")
            print("✅ Template updated.")
            commands = new_lines
            break
        elif action == 'd':
            try:
                delete_indices = input("🧹 Enter line numbers to delete (comma-separated): ")
                to_delete = {int(x.strip()) for x in delete_indices.split(",")}
                commands = [cmd for i, cmd in enumerate(commands, 1) if i not in to_delete]
                with open(template_file, "w") as f:
                    f.write("\n".join(commands) + "\n")
                print("✅ Selected lines deleted.")
                break
            except Exception as e:
                print(f"❌ Error deleting lines: {e}")
        elif action == 'e':
            print("🤖 Asking LLM for a plain summary of the template commands...\n")
            try:
                summary_prompt = f"""
Explain the following Python code to a non-technical person in 1–3 very short lines. Use simple, friendly language.

Code:
{chr(10).join(commands)}
"""
                summary_response = ollama.chat(
                    model='mistral',
                    messages=[{'role': 'user', 'content': summary_prompt}]
                )
                one_liner = summary_response['message']['content'].strip()
                print(f"🧾 In short: {one_liner}")
            except Exception as e:
                print(f"❌ Failed to query LLM for explanation: {e}")
        elif action == '':
            break
        else:
            print("⚠️ Invalid option. Try again.")

    # Execute the final commands
    local_vars = {'df': df, 'np': np, 'pd': pd, 'os': os}
    for cmd in commands:
        print(f"⚙️ Running: {cmd}")
        try:
            exec(cmd, {}, local_vars)
        except Exception as e:
            print(f"❌ Error while executing: {cmd}\n{e}")
            sys.exit(1)

    df = local_vars['df']
    print("✅ All commands executed.")

    # Save modified DataFrame
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = os.path.splitext(file_path)[0] + f"_modified_{timestamp}.xlsx"
    try:
        df.to_excel(output_path, index=False)
        print(f"\n✅ Modified Excel file saved as '{output_path}'")
    except Exception as e:
        print(f"❌ Failed to save Excel file: {e}")

# === Else: New Template Mode ===
else:
    print(f"\n🆕 Template file '{template_file}' not found. Creating a new one...")
    open(template_file, "w").close()

    while True:
        print("\n📋 Current Columns:", list(df.columns))
        print(df.head(3))
        data_dict = df.head(10).to_dict(orient='records')
        yaml_data = yaml.dump(data_dict, sort_keys=False)

        instruction = input("\n🧠 Enter your instruction (or type ':q' to quit):\n")
        if instruction.strip().lower() == ":q":
            break

        prompt = f"""
You are a Python expert. A user will give you a sample of Excel data (in YAML format) and ask for a transformation.

The YAML data is for context only. Do NOT try to parse it. It just shows what the DataFrame (`df`) looks like.

You must generate valid Python code that modifies an existing Pandas DataFrame named `df` **in-place**.

⚠️ Do not use 'yaml', 'your_yaml_string', or try to load any data — assume the `df` already exists and contains the data shown.

Examples:
- Add column "Total" = Price + Tax
- Fill missing values in 'Quantity' with 0
- Rename column 'cust_name' to 'CustomerName'
- Filter out rows where Age < 18

---

📄 YAML sample data:
{yaml_data}

📝 User's instruction:
{instruction}

👨‍💻 Your response: ONLY Python code that modifies `df` in place, no markdown, no explanations.
"""

        try:
            response = ollama.chat(
                model='mistral',
                messages=[{'role': 'user', 'content': prompt}]
            )
        except Exception as e:
            print(f"❌ Error querying LLM: {e}")
            continue

        raw_response = response['message']['content']
        code_match = re.search(r"```(?:python)?(.*?)```", raw_response, re.DOTALL)
        raw_code = code_match.group(1) if code_match else raw_response
        generated_code = textwrap.dedent(raw_code).strip()

        # Normalize and split lines
        lines = [line.strip() for line in generated_code.splitlines() if line.strip()]

        print("\n📦 Generated Python Code:")
        for line in lines:
            print(line)

        # Save each line to template file
        with open(template_file, "a") as f:
            for line in lines:
                f.write(line + "\n")
        print("💾 Command saved to template file.")

        # Execute each line
        local_vars = {'df': df, 'np': np, 'pd': pd, 'os': os}
        success = True
        for line in lines:
            print(f"⚙️ Executing: {line}")
            try:
                exec(line, {}, local_vars)
            except Exception as e:
                print(f"❌ Error executing: {line}\n{e}")
                success = False
                break

        if success:
            df = local_vars['df']
            print("✅ Command executed.")
            print(df.head())

    # Save modified DataFrame
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = os.path.splitext(file_path)[0] + f"_modified_{timestamp}.xlsx"
    try:
        df.to_excel(output_path, index=False)
        print(f"\n✅ All changes saved to '{output_path}'")
    except Exception as e:
        print(f"❌ Failed to save Excel file: {e}")
