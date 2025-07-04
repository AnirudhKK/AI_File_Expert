import streamlit as st
import pandas as pd
import yaml
import ollama
import os
import re
import textwrap
import numpy as np
from datetime import datetime
import ast

# === Streamlit Config ===
st.set_page_config(page_title="Excel Editor", layout="wide")
st.title("📊 Excel Editor using LLM + Templates")

# === Session State Initialization ===
if 'df' not in st.session_state:
    st.session_state.df = None
if 'template_file' not in st.session_state:
    st.session_state.template_file = ""
if 'preview_rows' not in st.session_state:
    st.session_state.preview_rows = 20
if 'template_loaded' not in st.session_state:
    st.session_state.template_loaded = False
if 'output_file' not in st.session_state:
    st.session_state.output_file = ""

# === File Upload ===
uploaded_file = st.file_uploader("📁 Upload Excel File", type=["xlsx"])
if uploaded_file:
    st.session_state.df = pd.read_excel(uploaded_file)
    st.success("✅ File loaded.")
    st.info("ℹ️ Now enter a template name below to continue.")

    # === User-Defined Preview Rows (Max 50) ===
    st.session_state.preview_rows = st.slider("🗃 Preview Rows", min_value=5, max_value=50, value=20, step=5)

    if st.button("👁 Show Preview"):
        st.dataframe(st.session_state.df.head(st.session_state.preview_rows))
        st.markdown("📌 **Available Columns:**")
        st.code(", ".join(st.session_state.df.columns))

    # === Template Name Input ===
    template_name = st.text_input("📄 Template Name", value="default_template")
    if template_name:
        st.markdown("✅ Template name accepted. Click below to load or create template.")

    template_file = re.sub(r'\s+', '_', template_name.strip().lower()) + ".txt"
    st.session_state.template_file = template_file

    # === Load or Create Template ===
    if st.button("📥 Load/Create Template"):
        if os.path.exists(template_file):
            with open(template_file, "r") as f:
                st.session_state.commands = [line.strip() for line in f if line.strip()]
            st.success(f"✅ Loaded existing template: `{template_file}`")
        else:
            open(template_file, "w").close()
            st.session_state.commands = []
            st.success(f"🆕 Created new template: `{template_file}`")
        st.session_state.template_loaded = True

    if st.session_state.get("template_loaded", False):
        commands = st.session_state.commands
        st.code("\n".join(commands), language="python")

        action = st.radio("⚙️ Manage Template", ["None", "Edit", "Delete Lines", "Explain"])

        if action == "Edit":
            new_text = st.text_area("📝 Edit template:", value="\n".join(commands))
            if st.button("💾 Save Template"):
                with open(template_file, "w") as f:
                    f.write(new_text.strip() + "\n")
                st.success(f"✅ Template `{template_file}` saved successfully.")
                st.session_state.commands = new_text.strip().splitlines()

        elif action == "Delete Lines":
            to_delete = st.multiselect("Select lines to delete", [f"{i+1}: {cmd}" for i, cmd in enumerate(commands)])
            if not to_delete:
                st.warning("⚠️ No lines selected for deletion.")
            else:
                indices = [int(line.split(":")[0]) - 1 for line in to_delete]
                commands = [cmd for i, cmd in enumerate(commands) if i not in indices]
                with open(template_file, "w") as f:
                    f.write("\n".join(commands) + "\n")
                st.success(f"✅ Selected lines deleted from `{template_file}`.")
                st.session_state.commands = commands
                st.code("\n".join(commands), language="python")

        elif action == "Explain":
            if not commands:
                st.warning("⚠️ This template is currently empty. Nothing to explain.")
            else:
                try:
                    summary_prompt = f"""
Explain the following Python code to a non-technical person in 1–3 very short lines. Use simple, friendly language.

Code:
{chr(10).join(commands)}
"""
                    with st.spinner("🧪 Generating explanation from LLM..."):
                        summary_response = ollama.chat(
                            model='mistral',
                            messages=[{'role': 'user', 'content': summary_prompt}]
                        )
                    st.info("🧾 " + summary_response['message']['content'].strip())
                except Exception as e:
                    st.error(f"❌ LLM error: {e}")

        if st.button("▶️ Run Template"):
            if not commands:
                st.warning("⚠️ No commands found in template. Please add or edit instructions.")
            else:
                with st.spinner("⚙️ Applying template to DataFrame..."):
                    local_vars = {'df': st.session_state.df.copy(), 'np': np, 'pd': pd, 'os': os}
                    try:
                        for cmd in commands:
                            if "df.append" in cmd:
                                st.warning("⚠️ Replacing deprecated df.append() with pd.concat().")
                                cmd = cmd.replace("df.append", "pd.concat([df,").replace("]", "], ignore_index=True)")
                            exec(cmd, {}, local_vars)
                        st.session_state.df = local_vars['df']
                        st.success(f"✅ Template `{template_file}` applied successfully.")
                        st.dataframe(st.session_state.df.head(st.session_state.preview_rows))

                        # Save Excel file after applying template
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        output_file = f"modified_{timestamp}.xlsx"
                        st.session_state.output_file = output_file
                        st.session_state.df.to_excel(output_file, index=False)
                        st.markdown(f"📁 Your file has been saved as: `{output_file}`")
                        with open(output_file, "rb") as f:
                            st.download_button("📥 Download Modified Excel", f, file_name=output_file)
                    except Exception as e:
                        st.error(f"❌ Template execution error:\n{e}")

    # === YAML Preview ===
    yaml_sample = yaml.dump(st.session_state.df.head(10).to_dict(orient='records'), sort_keys=False)
    with st.expander("📄 YAML Preview of Excel Sample"):
        st.code(yaml_sample, language="yaml")

    # === Natural Language Instruction Input ===
    instruction = st.text_input("🧠 Describe the transformation (e.g. 'add Total = Price + Tax'):")
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
            with st.spinner("🧪 Calling LLM for code generation..."):
                response = ollama.chat(
                    model='mistral',
                    messages=[{'role': 'user', 'content': prompt}]
                )
            raw_code = response['message']['content']

            code_match = re.search(r"```(?:python)?(.*?)```", raw_code, re.DOTALL)
            code_block = code_match.group(1) if code_match else raw_code

            dedented = textwrap.dedent(code_block)
            lines = [
                re.sub(r'^\s+', '', line).rstrip()
                for line in dedented.splitlines()
                if line.strip() and not line.strip().startswith("#")
            ]
            clean_code = "\n".join(lines)

            try:
                ast.parse(clean_code)
            except SyntaxError as e:
                st.error(f"❌ Syntax error in generated code:\n{e}")
                st.stop()

            st.code(clean_code, language="python")

            if st.button("🚀 Apply Transformation"):
                with st.spinner("🔧 Applying your transformation..."):
                    with open(template_file, "a") as f:
                        f.write(clean_code + "\n")
                    local_vars = {'df': st.session_state.df.copy(), 'np': np, 'pd': pd, 'os': os}
                    try:
                        exec(clean_code, {}, local_vars)
                        st.session_state.df = local_vars['df']
                        st.success(f"✅ Transformation applied and saved to `{template_file}`.")
                        st.dataframe(st.session_state.df.head(st.session_state.preview_rows))

                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        output_file = f"modified_{timestamp}.xlsx"
                        st.session_state.output_file = output_file
                        st.session_state.df.to_excel(output_file, index=False)
                        st.markdown(f"📁 Your file has been saved as: `{output_file}`")
                        with open(output_file, "rb") as f:
                            st.download_button("📥 Download Modified Excel", f, file_name=output_file)
                    except Exception as e:
                        st.error(f"❌ Execution error:\n{e}")
        except Exception as e:
            st.error(f"❌ LLM call failed:\n{e}")
