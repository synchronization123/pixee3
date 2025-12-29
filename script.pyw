# script.pyw
# Meeting Notes Maker – FINAL FIX
# Transcript-only truth
# Generates:
# 1) meeting_notes_YYYY-MM-DD.html
# 2) summary.html (AI-extracted project table)
# OpenRouter API | GUI-safe | Python 3.12

import os
import json
import datetime
import re
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from docx import Document
import urllib.request

# ===================== PATH CONFIG =====================
BASE_CODEX = r"C:\Users\Admin\Desktop\codex"
NOTES_MD = os.path.join(BASE_CODEX, "notes.md")
CONFIG_JSON = os.path.join(BASE_CODEX, "config.json")

# ===================== LOAD CONFIG =====================
if not os.path.exists(CONFIG_JSON):
    messagebox.showerror(
        "Missing config.json",
        f"config.json not found at:\n{CONFIG_JSON}"
    )
    raise SystemExit

with open(CONFIG_JSON, "r", encoding="utf-8") as cfg_file:
    CONFIG = json.load(cfg_file)

API_KEY = CONFIG.get("api_key")
MODEL = CONFIG.get("model")
API_URL = CONFIG.get("api_url")

if not API_KEY or not MODEL or not API_URL:
    messagebox.showerror(
        "Invalid config.json",
        "config.json must contain:\n"
        "- api_key\n- model\n- api_url"
    )
    raise SystemExit

# ===================== DOCX READER =====================
def read_docx(path):
    document = Document(path)
    lines = []
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if text:
            lines.append(text)
    return "\n".join(lines)

# ===================== NOTES TEMPLATE =====================
def read_notes_template():
    if not os.path.exists(NOTES_MD):
        messagebox.showerror("Missing notes.md", NOTES_MD)
        return None
    with open(NOTES_MD, "r", encoding="utf-8") as file:
        return file.read()

# ===================== HTML NORMALIZER =====================
def normalize_html(content):
    replacements = [
        (r'</section>', '</section>\n\n'),
        (r'<section>', '\n<section>\n'),
        (r'</h2>', '</h2>\n'),
        (r'</p>', '</p>\n'),
        (r'</ul>', '</ul>\n'),
        (r'</li>', '</li>\n'),
        (r'</tr>', '</tr>\n'),
    ]
    for pattern, repl in replacements:
        content = re.sub(pattern, repl, content, flags=re.IGNORECASE)
    return content.strip()

# ===================== OPENROUTER CALL =====================
def call_openrouter(system_prompt, user_prompt):
    payload = {
        "model": MODEL,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": 0.0,
        "max_tokens": CONFIG.get("max_tokens", 3000)
    }

    request = urllib.request.Request(
        API_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {API_KEY}",
            "Content-Type": "application/json"
        }
    )

    try:
        with urllib.request.urlopen(request, timeout=CONFIG.get("timeout", 60)) as response:
            data = json.loads(response.read().decode("utf-8"))
            return normalize_html(data["choices"][0]["message"]["content"])
    except Exception as e:
        messagebox.showerror("OpenRouter Error", str(e))
        return None

# ===================== MEETING NOTES =====================
def generate_meeting_notes(transcript, template):
    today = datetime.date.today().isoformat()

    system_prompt = (
        "You generate professional Minutes of Meeting.\n"
        "The transcript is the ONLY source of truth.\n"
        "Do NOT invent content.\n"
        "If something is missing, write 'Not discussed in the meeting.'\n"
        "Output ONLY HTML body using <section>, <h2>, <p>, <ul>, <li>."
    )

    user_prompt = f"""
Meeting Date: {today}

### TEMPLATE ###
{template}

### TRANSCRIPT ###
{transcript}
"""

    return call_openrouter(system_prompt, user_prompt)

# ===================== SUMMARY TABLE (FIXED) =====================
def generate_summary_html(transcript):
    today = datetime.date.today().isoformat()
    output_file = os.path.join(BASE_CODEX, "summary.html")

    system_prompt = (
        "You extract project information STRICTLY from the transcript.\n"
        "Rules:\n"
        "- Only include projects explicitly mentioned.\n"
        "- Do NOT invent projects or details.\n"
        "- If a column is not mentioned, write 'Not discussed in the meeting.'\n"
        "- If NO projects are discussed, create ONE row stating 'No projects discussed'.\n"
        "- Output ONLY valid HTML <table> body (no html/head/body tags).\n"
        "- Use <tr>, <td>, <ul>, <li> where appropriate."
    )

    user_prompt = f"""
Generate a project summary table with the following columns:

1. Sr. No.
2. Project Name
3. Owners from Team
4. Value Add
5. Project Category
6. Expected Completion
7. Status
8. Start Date
9. Completed Date
10. Milestones
11. Challenges
12. Tracker
13. Action Items

### TRANSCRIPT ###
{transcript}
"""

    table_body = call_openrouter(system_prompt, user_prompt)
    if table_body is None:
        return None

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Project Summary</title>
<style>
body {{
    font-family: Segoe UI, Arial, sans-serif;
    margin: 40px;
}}
table {{
    border-collapse: collapse;
    width: 100%;
    table-layout: fixed;
}}
th, td {{
    border: 1px solid #cbd5e1;
    padding: 8px;
    vertical-align: top;
    word-wrap: break-word;
}}
th {{
    background-color: #f1f5f9;
}}
</style>
</head>
<body>

<h1>Project Summary – {today}</h1>

<table>
<thead>
<tr>
<th>Sr. No.</th>
<th>Project Name</th>
<th>Owners from Team</th>
<th>Value Add</th>
<th>Project Category</th>
<th>Expected Completion</th>
<th>Status</th>
<th>Start Date</th>
<th>Completed Date</th>
<th>Milestones</th>
<th>Challenges</th>
<th>Tracker</th>
<th>Action Items</th>
</tr>
</thead>
<tbody>
{table_body}
</tbody>
</table>

</body>
</html>
"""

    with open(output_file, "w", encoding="utf-8") as file:
        file.write(html)

    return output_file

# ===================== HTML WRITER =====================
def write_meeting_html(content):
    today = datetime.date.today().isoformat()
    output_file = os.path.join(BASE_CODEX, f"meeting_notes_{today}.html")

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Meeting Notes</title>
<style>
body {{
    font-family: Segoe UI, Arial, sans-serif;
    margin: 40px;
    line-height: 1.7;
}}
section {{
    margin-bottom: 40px;
}}
h2 {{
    border-bottom: 1px solid #cbd5e1;
    padding-bottom: 6px;
}}
</style>
</head>
<body>

<h1>Meeting Notes – {today}</h1>

{content}

</body>
</html>
"""

    with open(output_file, "w", encoding="utf-8") as file:
        file.write(html)

    return output_file

# ===================== GUI =====================
class MeetingNotesApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Meeting Notes Maker (Notes + Project Summary)")

        screen_w = root.winfo_screenwidth()
        screen_h = root.winfo_screenheight()
        self.root.geometry(f"{screen_w}x{screen_h}")
        self.root.state("zoomed")

        self.transcript = ""

        container = tk.Frame(root)
        container.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            container,
            text="AI Provider: OPENROUTER (TRANSCRIPT-ONLY)",
            font=("Segoe UI", 11, "bold")
        ).pack(pady=10)

        tk.Button(
            container,
            text="Import DOCX Transcription",
            width=50,
            command=self.import_docx
        ).pack(pady=10)

        self.generate_btn = tk.Button(
            container,
            text="Generate Meeting Notes & Summary",
            width=50,
            state="disabled",
            command=self.generate_all
        )
        self.generate_btn.pack(pady=10)

        self.preview = scrolledtext.ScrolledText(
            container,
            wrap=tk.WORD,
            font=("Consolas", 10)
        )
        self.preview.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def import_docx(self):
        path = filedialog.askopenfilename(
            filetypes=[("Word Documents", "*.docx")]
        )
        if not path:
            return

        self.transcript = read_docx(path)
        self.preview.delete("1.0", tk.END)
        self.preview.insert(tk.END, self.transcript)
        self.generate_btn.config(state="normal")

    def generate_all(self):
        template = read_notes_template()
        if template is None:
            return

        self.preview.delete("1.0", tk.END)
        self.preview.insert(tk.END, "Generating meeting notes and project summary...\n\n")

        notes_html = generate_meeting_notes(self.transcript, template)
        if notes_html is None:
            return

        notes_file = write_meeting_html(notes_html)
        summary_file = generate_summary_html(self.transcript)

        messagebox.showinfo(
            "Success",
            f"Files generated successfully:\n\n{notes_file}\n{summary_file}"
        )

# ===================== MAIN =====================
if __name__ == "__main__":
    root = tk.Tk()
    app = MeetingNotesApp(root)
    root.mainloop()
