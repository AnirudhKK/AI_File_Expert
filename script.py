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

# === Convert DataFrame to YAML ===
data_dict = df.head(10).to_dict(orient='records')  # sample first 10 rows for context
yaml_data = yaml.dump(data_dict, sort_keys=False)

# === Get user instruction ===
instruction = input("\n🧠 Describe the changes you want to make to the Excel data:\n")

# === Construct prompt for Mistral ===
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

# === Query Mistral model via Ollama ===
response = ollama.chat(
    model='mistral',
    messages=[{'role': 'user', 'content': prompt}]
)

# Extract only the Python code (remove any markdown formatting)
raw_response = response['message']['content']
code_match = re.search(r"```(?:python)?(.*?)```", raw_response, re.DOTALL)

if code_match:
    generated_code = code_match.group(1).strip()
else:
    generated_code = raw_response.strip()

print("\n📦 Generated Python Code:\n", generated_code)

# === Execute the generated code safely ===
try:
    exec(generated_code, {'df': df})
except Exception as e:
    print("❌ Error executing the generated code:", e)
    sys.exit(1)

# === Save modified Excel ===
output_path = os.path.splitext(file_path)[0] + "_modified.xlsx"
df.to_excel(output_path, index=False)

print(f"\n✅ Changes applied and saved to '{output_path}'")