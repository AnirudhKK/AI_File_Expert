import pandas as pd
import ollama
import yaml
import sys
import os
import re
import textwrap

# === Ask user for Excel file path ===
file_path = input("🔍 Enter the path to your Excel file: ").strip()

if not os.path.exists(file_path):
    print("❌ File not found. Please check the path.")
    sys.exit(1)

# === Load Excel data ===
try:
    df = pd.read_excel(file_path)
except Exception as e:
    print(f"❌ Failed to load Excel file: {e}")
    sys.exit(1)

# === Interactive loop ===
while True:
    print("\n📋 Current Columns:", list(df.columns))
    print(df.head(3))  # Optional: preview

    # === Sample YAML for LLM context ===
    data_dict = df.head(10).to_dict(orient='records')
    yaml_data = yaml.dump(data_dict, sort_keys=False)

    # === Get user instruction ===
    instruction = input("\n🧠 Enter your instruction (or type ':q' to quit):\n")
    if instruction.strip().lower() == ":q":
        break

    # === Build LLM prompt ===
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

    # === Query Mistral via Ollama ===
    try:
        response = ollama.chat(
            model='mistral',
            messages=[{'role': 'user', 'content': prompt}]
        )
    except Exception as e:
        print(f"❌ Error querying LLM: {e}")
        continue

    # === Extract and normalize generated code ===
    import textwrap
    raw_response = response['message']['content']
    code_match = re.search(r"```(?:python)?(.*?)```", raw_response, re.DOTALL)

    if code_match:
        raw_code = code_match.group(1)
    else:
        raw_code = raw_response

    # 🔥 Critical fix — normalize all lines by dedenting
    lines = raw_code.splitlines()
    dedented = textwrap.dedent("\n".join(lines)).strip()
    generated_code = dedented


    print("\n📦 Generated Python Code:\n", generated_code)

    # === Validate referenced columns (read-only) ===
    col_names = set(df.columns.str.lower())
    written_cols = set(re.findall(r"df\[['\"]([a-zA-Z0-9_ ]+)['\"]\]\s*=", generated_code))
    all_cols = set(re.findall(r"df\[['\"]([a-zA-Z0-9_ ]+)['\"]\]", generated_code))
    read_cols = all_cols - written_cols

    unknown_cols = [col for col in read_cols if col.lower() not in col_names]
    if unknown_cols:
        print(f"⚠️ Warning: These columns are not found in your Excel data: {unknown_cols}")
        confirm = input("Do you still want to run this code? (y/n): ").strip().lower()
        if confirm != 'y':
            print("⏭️ Skipped this change.")
            continue  # ← now valid, inside the loop

    # === Try to execute the generated code ===
    # === Try to execute the generated code ===
    try:
        local_vars = {'df': df}
        exec(generated_code, {}, local_vars)
        df = local_vars['df']
        print("✅ Change applied.")
        print("\n📊 Updated DataFrame Preview:\n", df.head())
    except Exception as e:
        print("❌ Error executing generated code:", e)

# === Save modified file ===
output_path = os.path.splitext(file_path)[0] + "_modified.xlsx"
try:
    df.to_excel(output_path, index=False)
    print(f"\n✅ All changes saved to '{output_path}'")
except Exception as e:
    print(f"❌ Failed to save Excel file: {e}")
