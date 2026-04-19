"""
app.py
Streamlit UI — thin layer that delegates all logic to modules.

Security measures applied here:
  - pd.read_excel uses engine="openpyxl" explicitly (no macro-capable engine).
  - Downloaded Excel is built in-memory (BytesIO) — no temp file left on disk.
  - _download_widget defined before first use.

Tabs:
  Tab 1 — Single File Editor  (LLM + template transforms on one file)
  Tab 2 — Compare & Merge     (multi-file diff/update: mark closed, add new)
"""
import io
from datetime import datetime

import pandas as pd
import streamlit as st
import yaml

from template_manager import (
    append_to_template,
    delete_template_lines,
    ensure_template_exists,
    get_template_filename,
    load_template,
    save_template,
)
from code_generator import explain_commands, generate_code
from df_executor import run_commands, run_single
from file_merger import MergeConfig, MergeResult, merge_files, validate_columns  # noqa: F401

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Excel Editor", layout="wide")
st.title("📊 Excel Editor using LLM + Templates")

# ── Session state ─────────────────────────────────────────────────────────────
if "df" not in st.session_state:
    st.session_state.df = None
if "template_file" not in st.session_state:
    st.session_state.template_file = ""
if "merge_result" not in st.session_state:
    st.session_state.merge_result = None


# ── Shared download helper ────────────────────────────────────────────────────
def _download_widget(df: pd.DataFrame, label: str = "📥 Download Excel") -> None:
    """Render a download button that streams the DataFrame as an xlsx in memory."""
    if df is None:
        return
    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl")
    buf.seek(0)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    st.download_button(
        label,
        data=buf,
        file_name=f"modified_{timestamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ── Tabs ──────────────────────────────────────────────────────────────────────
tab_single, tab_merge = st.tabs(["📝 Single File Editor", "🔀 Compare & Merge"])


# ═════════════════════════════════════════════════════════════════════════════
# TAB 1 — Single File Editor
# Wrapped in a function so `return` can be used instead of `st.stop()`.
# st.stop() is GLOBAL in Streamlit — it would kill Tab 2 as well.
# ═════════════════════════════════════════════════════════════════════════════
def _render_single_file_tab() -> None:
    # ── File upload ───────────────────────────────────────────────────────────
    uploaded_file = st.file_uploader("📁 Upload Excel File", type=["xlsx"], key="single_upload")
    if not uploaded_file:
        st.info("Upload an Excel file to get started.")
        return

    # Pin engine to openpyxl — prevents xlrd from executing macros/external links
    st.session_state.df = pd.read_excel(uploaded_file, engine="openpyxl")
    st.success("✅ File loaded.")
    st.dataframe(st.session_state.df.head())

    # ── Template setup ────────────────────────────────────────────────────────
    template_name = st.text_input("📄 Template Name", value="default_template")
    template_file = get_template_filename(template_name)
    st.session_state.template_file = template_file

    already_existed = ensure_template_exists(template_file)
    commands = load_template(template_file)

    # ── Existing template UI ──────────────────────────────────────────────────
    if already_existed and commands:
        st.markdown(f"📂 Using existing template: `{template_file}`")
        st.code("\n".join(commands), language="python")

        action = st.radio("⚙️ Manage Template", ["None", "Edit", "Delete Lines", "Explain"])

        if action == "Edit":
            new_text = st.text_area("📝 Edit template:", value="\n".join(commands))
            if st.button("💾 Save Template"):
                updated = new_text.strip().splitlines()
                save_template(template_file, updated)
                commands = updated
                st.success("✅ Template saved.")

        elif action == "Delete Lines":
            labeled = [f"{i+1}: {cmd}" for i, cmd in enumerate(commands)]
            to_delete = st.multiselect("Select lines to delete", labeled)
            indices = [int(lbl.split(":")[0]) - 1 for lbl in to_delete]
            if indices:
                commands = delete_template_lines(template_file, commands, indices)
                st.success("✅ Lines deleted.")

        elif action == "Explain":
            try:
                explanation = explain_commands(commands)
                st.info("🧾 " + explanation)
            except Exception as e:
                st.error(f"❌ LLM error: {e}")

        if st.button("▶️ Run Template"):
            try:
                result_df, warnings = run_commands(st.session_state.df, commands)
                for w in warnings:
                    st.warning(f"⚠️ {w}")
                st.session_state.df = result_df
                st.success("✅ Template applied.")
                st.dataframe(st.session_state.df.head())
            except (ValueError, RuntimeError) as e:
                st.error(f"❌ {e}")

    else:
        st.markdown(f"🆕 Creating new template: `{template_file}`")

    # ── YAML preview ──────────────────────────────────────────────────────────
    yaml_sample = yaml.dump(
        st.session_state.df.head(10).to_dict(orient="records"), sort_keys=False
    )
    with st.expander("📄 YAML Preview of Excel Sample"):
        st.code(yaml_sample, language="yaml")

    # ── Download (always visible once a file is loaded) ───────────────────────
    _download_widget(st.session_state.df)

    # ── LLM transformation ────────────────────────────────────────────────────
    instruction = st.text_input("🧠 Describe the transformation (e.g. 'add Total = Price + Tax'):")
    if not instruction:
        return

    try:
        clean_code, error = generate_code(yaml_sample, instruction)
    except Exception as e:
        st.error(f"❌ LLM call failed:\n{e}")
        return

    if error:
        st.error(f"❌ Code rejected: {error}")
        return

    st.code(clean_code, language="python")

    if st.button("🚀 Apply Transformation"):
        try:
            result_df, warnings = run_single(st.session_state.df, clean_code)
            for w in warnings:
                st.warning(f"⚠️ {w}")
            append_to_template(template_file, clean_code)
            st.session_state.df = result_df
            st.success("✅ Transformation applied and saved.")
            st.dataframe(st.session_state.df.head())
            _download_widget(st.session_state.df)
        except (ValueError, RuntimeError) as e:
            st.error(f"❌ {e}")


with tab_single:
    _render_single_file_tab()


# ═════════════════════════════════════════════════════════════════════════════
# TAB 2 — Compare & Merge
# ═════════════════════════════════════════════════════════════════════════════
def _render_merge_tab() -> None:
    st.subheader("🔀 Compare & Merge Two Excel Files")
    st.markdown(
        "Upload your **master file** (full history: open + closed rows) and "
        "the **new export** (only currently open items). The merge will:\n"
        "- 🔒 Mark rows missing from the new export as **Closed**\n"
        "- ➕ Append rows in the new export that don't exist yet in the master"
    )

    col_master, col_new = st.columns(2)
    with col_master:
        st.markdown("#### 📋 Master File")
        master_upload = st.file_uploader(
            "Full history (open + closed)", type=["xlsx"], key="master_upload"
        )
    with col_new:
        st.markdown("#### 📤 New Export")
        new_upload = st.file_uploader(
            "Latest snapshot (open items only)", type=["xlsx"], key="new_upload"
        )

    if not master_upload or not new_upload:
        st.info("Upload both files above to configure the merge.")
        return

    # ── Load both files ───────────────────────────────────────────────────────
    master_df = pd.read_excel(master_upload, engine="openpyxl")
    new_df    = pd.read_excel(new_upload,    engine="openpyxl")

    master_cols = list(master_df.columns)
    common_cols = [c for c in master_cols if c in new_df.columns]

    with st.expander("👀 Preview master file"):
        st.dataframe(master_df.head(10), use_container_width=True)
    with st.expander("👀 Preview new export"):
        st.dataframe(new_df.head(10), use_container_width=True)

    st.divider()
    st.markdown("### ⚙️ Merge Configuration")

    col_cfg1, col_cfg2 = st.columns(2)

    with col_cfg1:
        # Single ID column — selectbox, not multiselect
        id_col = st.selectbox(
            "🔑 ID column — uniquely identifies each action item",
            options=common_cols,
            index=0,
            help="Must exist in both files. E.g. 'Action Item ID', 'Ticket No'.",
        )

    with col_cfg2:
        status_col = st.selectbox(
            "📌 Status column — holds Open/Closed in the master",
            options=master_cols,
            index=master_cols.index("Status") if "Status" in master_cols else 0,
        )

    col_ov, col_cv = st.columns(2)
    with col_ov:
        open_value   = st.text_input("✅ Value meaning Open",   value="Open")
    with col_cv:
        closed_value = st.text_input("🔒 Value to write when closing", value="Closed")

    st.divider()

    # ── Validate ──────────────────────────────────────────────────────────────
    cfg = MergeConfig(
        id_col=id_col,
        status_col=status_col,
        closed_value=closed_value,
        open_value=open_value,
    )
    for err in validate_columns(master_df, new_df, cfg):
        st.warning(f"⚠️ {err}")

    # ── Preview ───────────────────────────────────────────────────────────────
    if st.button("🔍 Preview Changes"):
        try:
            result: MergeResult = merge_files(master_df, new_df, cfg)
            st.session_state.merge_result = result

            st.markdown("### 📊 Change Summary")
            st.markdown(result.summary)

            if result.closed_ids:
                with st.expander(f"🔒 {len(result.closed_ids)} item(s) to be marked Closed"):
                    st.dataframe(
                        pd.DataFrame({id_col: result.closed_ids}),
                        use_container_width=True,
                    )
            if result.added_ids:
                with st.expander(f"➕ {len(result.added_ids)} new item(s) to be added"):
                    st.dataframe(
                        pd.DataFrame({id_col: result.added_ids}),
                        use_container_width=True,
                    )

            st.info("Review above, then click **Apply & Download** to get the updated file.")

        except Exception as e:
            st.error(f"❌ Merge failed: {e}")

    # ── Apply & download ──────────────────────────────────────────────────────
    if st.session_state.merge_result is not None:
        if st.button("✅ Apply & Download Updated Master"):
            _download_widget(
                st.session_state.merge_result.df,
                label="📥 Download Updated Master",
            )
            st.success("Updated master file ready for download 👆")
            st.dataframe(
                st.session_state.merge_result.df,
                use_container_width=True,
            )


with tab_merge:
    _render_merge_tab()
