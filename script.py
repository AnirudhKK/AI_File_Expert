import streamlit as st
import pandas as pd
import yaml
import ollama
import os
import re
import textwrap
import numpy as np
from datetime import datetime

st.set_page_config(page_title="Excel Editor with LLM", layout="wide")
st.title("📊 Excel Editor (LLM + YAML + Templates)")

# === Initialize session state ===
if 'df' not in st.session_state:
    st.session_state.df = None
if 'template_file' not in st.session_state:
    st.session_state.template_file = ""

# === Upload Excel File ===
uploaded_file = st.file_uploader("📁 Upload an Excel file", type=["xlsx"])

if uploaded_file:
    st.session_state.df = pd.read_excel(uploaded_file)
    st.success("✅ File loaded successfully.")
    st.dataframe(st.session_state.df.head())

    # === Template Name ===
    template_name = st.text_input("📄 Template name (e.g., sales_cleanup):", value="default_template")
    template_file = re.sub(r'\s+', '_', template_name.strip().lower()) + ".txt"
    st.session_state.template_file = template_file

    # === Load Existing Template ===
    if os.path.exists(template_file):
        st.markdown(f"📂 Found template: `{template_file}`")

        with open(template_file, "r") as f:
            commands = [line.strip() for line in f if line.strip()]
        if commands:
            st.code("\n".join(commands), language="python")

            action = st.radio("Do you want to modify this template?", options=["No", "Edit", "Delete Lines", "Explain"])

            if action == "Edit":
                new_text = st.text_area("📝 Edit the template:", value="\n".join(commands))
                if st.button("💾 Save Template"):
                    with open(template_file, "w") as f:
                        f.write(new_text.strip() + "\n")
                    st.success("✅ Template updated.")
                    commands = new_text.strip().splitlines()

            elif action == "Delete Lines":
                to_delete = st.multiselect("Select lines to delete", [f"{i+1}: {cmd}" for i, cmd in enumerate(commands)])
                indices = [int(t.split(":")[0]) - 1 for t in to_delete]
                commands = [cmd for i, cmd in enumerate(commands) if i not in indices]
                with open(template_file, "w") as f:
                    f.write("\n".join(commands) + "\n")
                st.success("✅ Selected lines deleted.")

            elif action == "Explain":
                try:
                    explain_prompt = f"""
Explain the following Python code to a non-technical person in 1–3 very short lines. Use simple, friendly language.

Code:
{chr(10).join(commands)}
"""
                    summary_response = ollama.chat(
                        model='mistral',
                        messages=[{'role': 'user', 'content': explain_prompt}]
                    )
                    explanation = summary_response['message']['content'].strip()
                    st.info(f"💬 Explanation:\n\n{explanation}")
                except Exception as e:
                    st.error(f"❌ Failed to query LLM: {e}")

            if st.button("▶️ Run Template"):
                local_vars = {'df': st.session_state.df.copy(), 'np': np, 'pd': pd, 'os': os}
                for cmd in commands:
                    try:
                        exec(cmd, {}, local_vars)
                    except Exception as e:
                        st.error(f"❌ Error running `{cmd}`: {e}")
                        st.stop()
                st.session_state.df = local_vars['df']
                st.success("✅ Template applied.")
                st.dataframe(st.session_state.df.head())

    else:
        st.markdown(f"🆕 Creating new template: `{template_file}`")
        open(template_file, "w").close()

    # === YAML Preview ===
    yaml_sample = yaml.dump(st.session_state.df.head(10).to_dict(orient='records'), sort_keys=False)
    with st.expander("📄 YAML Preview (top 10 rows)"):
        st.code(yaml_sample, language="yaml")

    # === Input Instruction ===
    instruction = st.text_input("💡 Describe what changes you'd like to make:")
    if instruction:
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
{yaml_sample}

📝 User's instruction:
{instruction}

👨‍💻 Your response: ONLY Python code that modifies `df` in place, no markdown, no explanations.
"""
        try:
            response = ollama.chat(
                model='mistral',
                messages=[{'role': 'user', 'content': prompt}]
            )
            raw_code = response['message']['content']
            code_match = re.search(r"```(?:python)?(.*?)```", raw_code, re.DOTALL)
            clean_code = textwrap.dedent(code_match.group(1) if code_match else raw_code).strip()
            lines = [line for line in clean_code.splitlines() if line.strip()]

            st.code("\n".join(lines), language="python")

            if st.button("🚀 Run & Save Command"):
                with open(template_file, "a") as f:
                    for line in lines:
                        f.write(line + "\n")

                local_vars = {'df': st.session_state.df.copy(), 'np': np, 'pd': pd, 'os': os}
                for line in lines:
                    try:
                        exec(line, {}, local_vars)
                    except Exception as e:
                        st.error(f"❌ Error executing: `{line}`\n{e}")
                        st.stop()

                st.session_state.df = local_vars['df']
                st.success("✅ Transformation applied and saved.")
                st.dataframe(st.session_state.df.head())

        except Exception as e:
            st.error(f"❌ LLM Query Error: {e}")

    # === Export Final Excel ===
    if st.session_state.df is not None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"modified_{timestamp}.xlsx"
        st.session_state.df.to_excel(output_file, index=False)
        with open(output_file, "rb") as f:
            st.download_button("📥 Download Final Excel", f, file_name=output_file)
