"""
code_generator.py
Handles LLM prompting, code extraction, cleaning, and validation.

Security measures:
  - User instruction is length-capped and injected inside a delimited fence
    so that prompt-injection attempts ("ignore above instructions…") are
    clearly separated from the trusted system prompt.
  - Generated code is security-checked by sandbox.check_code() before being
    returned to the caller; callers must not exec() code that fails this check.
  - yaml_sample length is also capped so a crafted Excel file cannot inflate
    the prompt to an unreasonable size.
"""
import re
import ast
import textwrap
import ollama

from sandbox import check_code

MODEL = "mistral"

# Hard limits on user-controlled input injected into the LLM prompt
MAX_INSTRUCTION_LEN = 500    # characters
MAX_YAML_SAMPLE_LEN = 8_000  # characters


def _sanitise_instruction(instruction: str) -> str:
    """Truncate and strip control characters from the user instruction."""
    instruction = instruction[:MAX_INSTRUCTION_LEN]
    # Remove ASCII control characters (except newline/tab which are harmless)
    instruction = re.sub(r'[\x00-\x08\x0b-\x1f\x7f]', '', instruction)
    return instruction.strip()


def build_prompt(yaml_sample: str, instruction: str) -> str:
    # Cap yaml_sample to avoid prompt bloat from a crafted Excel file
    if len(yaml_sample) > MAX_YAML_SAMPLE_LEN:
        yaml_sample = yaml_sample[:MAX_YAML_SAMPLE_LEN] + "\n# ... (truncated)"

    safe_instruction = _sanitise_instruction(instruction)

    return f"""
You are a Python expert. A user will give you a sample of Excel data (in YAML \
format) and ask for a transformation.

The YAML data is for context only. Do NOT try to parse it. It shows what the \
DataFrame (`df`) looks like.

Generate valid Python code that modifies an existing Pandas DataFrame named \
`df` in-place.

RULES — you MUST follow all of them:
- Do NOT import any module (no import os, sys, subprocess, etc.)
- Do NOT use open(), eval(), exec(), compile(), or any file I/O
- Do NOT access __builtins__, __class__, __mro__, or any dunder attributes
- Only use `df`, `pd`, and `np` — nothing else
- Return ONLY Python code, no markdown fences, no explanations

Examples of acceptable operations:
- df["Total"] = df["Price"] + df["Tax"]
- df["Quantity"] = df["Quantity"].fillna(0)
- df.rename(columns={{"cust_name": "CustomerName"}}, inplace=True)
- df = df[df["Age"] >= 18]

---

📄 YAML sample data:
{yaml_sample}

--- USER INSTRUCTION START ---
{safe_instruction}
--- USER INSTRUCTION END ---

Respond with ONLY the Python code.
"""


def build_explain_prompt(commands: list[str]) -> str:
    # Cap total size of code being explained
    code_text = "\n".join(commands)[:3_000]
    return (
        "Explain the following Python code to a non-technical person in "
        "1–3 very short lines. Use simple, friendly language.\n\n"
        f"Code:\n{code_text}"
    )


def call_llm(prompt: str) -> str:
    response = ollama.chat(
        model=MODEL,
        messages=[{"role": "user", "content": prompt}]
    )
    return response["message"]["content"]


def extract_code(raw_response: str) -> str:
    """Extract and clean code from LLM response."""
    code_match = re.search(r"```(?:python)?(.*?)```", raw_response, re.DOTALL)
    code_block = code_match.group(1) if code_match else raw_response

    dedented = textwrap.dedent(code_block)
    lines = [
        re.sub(r'^\s+', '', line).rstrip()
        for line in dedented.splitlines()
        if line.strip() and not line.strip().startswith("#")
    ]
    return "\n".join(lines)


def validate_syntax(code: str) -> tuple[bool, str]:
    """Returns (is_valid, error_message)."""
    try:
        ast.parse(code)
        return True, ""
    except SyntaxError as e:
        return False, str(e)


def generate_code(yaml_sample: str, instruction: str) -> tuple[str, str]:
    """
    Full pipeline: prompt → LLM → extract → syntax-check → security-check.
    Returns (clean_code, error_message).
    error_message is '' when the code is safe and syntactically valid.
    """
    prompt = build_prompt(yaml_sample, instruction)
    raw = call_llm(prompt)
    clean = extract_code(raw)

    # 1. Syntax check
    valid, err = validate_syntax(clean)
    if not valid:
        return clean, f"Syntax error: {err}"

    # 2. Security check — catches dangerous patterns before exec()
    safe, reason = check_code(clean)
    if not safe:
        return clean, reason

    return clean, ""


def explain_commands(commands: list[str]) -> str:
    prompt = build_explain_prompt(commands)
    return call_llm(prompt).strip()
