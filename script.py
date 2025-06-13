import pandas as pd
import ollama
import yaml
import sys
import os
import re

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

# === Begin interactive modification loop ===
while True:
    # === Show YAML sample of current df ===
    data_dict = df.head(10).to_dict(orient='records')
    yaml_data = yaml.dump(data_dict, sort_keys=False)

    # === Ask user for instruction ===
    instruction = input("\n🧠 Describe the changes you want to make (or type ':q' to quit):\n")
    if instruction.strip().lower() == ":q":
        break

    # === Construct LLM prompt ===
    prompt = f"""
You are a Python expert. A user will give you a sample of Excel data (in YAML format) and ask for a transformation.

The YAML data is for context only. Do NOT try to parse it. It just shows what the DataFrame (`df`) looks like.

You must generate valid Python code that modifies an existing Pandas DataFrame named `df` **in-place**.

⚠️ Do not use 'yaml', 'your_yaml_string', or try to load any data — assume the `df` already exists and contains the data shown.

---

📄 YAML sample data:
{yaml_data}

📝 User's instruction:
{instruction}

👨‍💻 Your response: ONLY Python code that modifies `df` in place, no markdown, no explanations.
"""

    # === Query Ollama ===
    try:
        response = ollama.chat(
            model='mistral',
            messages=[{'role': 'user', 'content': prompt}]
        )
    except Exception as e:
        print(f"❌ Error communicating with LLM: {e}")
        continue

    # === Extract Python code ===
    raw_response = response['message']['content']
    code_match = re.search(r"```(?:python)?(.*?)```", raw_response, re.DOTALL)
    generated_code = code_match.group(1).strip() if code_match else raw_response.strip()

    print("\n📦 Generated Python Code:\n", generated_code)

    # === Execute code with safety ===
    try:
        exec(generated_code, {'df': df})
    except Exception as e:
        print("❌ Error executing the generated code:", e)
        continue

    print("✅ Change applied successfully.")

# === Save modified Excel ===
output_path = os.path.splitext(file_path)[0] + "_modified.xlsx"
try:
    df.to_excel(output_path, index=False)
    print(f"\n✅ All changes saved to '{output_path}'")
except Exception as e:
    print(f"❌ Failed to save Excel file: {e}")
