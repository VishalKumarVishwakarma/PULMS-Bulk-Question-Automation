# PULMS Bulk Question Uploader

Automates bulk creation of quiz questions on the **PULMS** exam platform
(Parul University LMS – an Angular/Ionic app at `elearning.paruluniversity.ac.in`).
It reads questions from an Excel file and fills the platform's **Add Question**
form for every row, then saves them to the exam.

The script attaches to an **already-open, already-logged-in Chrome tab** via
remote debugging, so you stay in full control of the browser and session.

---

## Features

- Reads a whole question bank from a single `.xlsx` file and inserts every row.
- Fills the rich-text **Question** and **Option** fields (Quill editors).
- Adds up to 4 options and marks the **correct** one.
- Selects the **Category** (auto-creates it via the modal if it doesn't exist).
- Sets the **Difficulty** (Easy / Medium / Hard).
- Turns the **Randomize Options** toggle ON for every question (configurable).
- Waits for the **Submit** button to become enabled, then submits.
- Auto-navigates back to a fresh form for the next question.
- Clicks **Save & Continue** at the end to persist everything to the server.
- **Per-row error recovery**: if one question fails, it recovers to a fresh
  form and continues with the next one (a single failure never stops the batch).
- Logs every row's success/failure and saves a screenshot on failure.
- Runs from **any folder** – finds `Question_Bank.xlsx` in the current directory.

---

## Requirements

- **Python 3.9+**
- **Google Chrome** (matching `chromedriver` is auto-downloaded by Selenium Manager)
- Python packages:

```bash
pip install selenium pandas openpyxl
```

---

## Excel format

The file must be named **`Question_Bank.xlsx`** (or pass a custom path as an
argument) with these **exact** column headers:

| Question | Option 1 | Option 2 | Option 3 | Option 4 | Correct Option | Difficulty Level |
|----------|----------|----------|----------|----------|----------------|------------------|
| What does 'Big Data' primarily refer to? | Small datasets stored in Excel | Extremely large and complex datasets | Only structured data in RDBMS | Data stored on a single computer | Option 2 | Easy |
| SQL stands for: | Structured Query Language | Simple Query Language | Sequential Query Language | Standard Query Language | Option 1 | Medium |

Notes:
- **One row = one question.** The header is row 1; questions start at row 2.
- **Correct Option** accepts any of these formats (auto-detected):
  - `Option 2` (the label), or
  - a number `1`–`4`, or
  - the exact option text.
- **Difficulty Level** must be `Easy`, `Medium`, or `Hard`.
- Fill all 4 options (minimum 2; blank options are ignored).

---

## Setup: start Chrome in debug mode

Close all Chrome windows, then launch Chrome with remote debugging:

```bat
chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\chrome-debug-profile"
```

Then, **manually**:
1. Log in to PULMS.
2. Open your exam and navigate to the **Question → Create Questions →
   Single Select** form (the blank "Add Question" screen).

Leave the browser on that form. The script attaches to it – it never opens
its own browser.

---

## Usage

Put `Question_Bank.xlsx` in any folder, open a terminal there, and run one of:

**1. Directly with Python**
```bat
python "C:\path\to\bulk_add_questions.py"
```

**2. Using the launcher (`insert_questions.bat`)**
```bat
insert_questions
```

**3. With an explicit Excel path (any name/location)**
```bat
python "C:\path\to\bulk_add_questions.py" "D:\my questions\bank.xlsx"
```

The script reads `Question_Bank.xlsx` from the **current folder** by default,
and writes its log + failure screenshots there.

> **Tip:** to run `insert_questions` from anywhere, add the script's folder to
> your Windows `PATH`.

While it runs, don't touch or reload that Chrome tab – the script adds every
question in one session and saves once at the end.

---

## Configuration

Edit the `CONFIG` block near the top of `bulk_add_questions.py`:

| Setting | Default | Description |
|---------|---------|-------------|
| `EXCEL_FILENAME` | `"Question_Bank.xlsx"` | File name the script looks for. |
| `DEBUGGER_ADDRESS` | `"127.0.0.1:9222"` | Chrome remote-debugging address. |
| `CATEGORY_NAME` | `"Weekly_exam"` | Category applied to every question. |
| `SUBCATEGORY_TITLE` | `None` | Sub Category (left untouched if `None`). |
| `RANDOMIZE_OPTIONS` | `True` | Turn the Randomize Options toggle ON. |
| `START_ROW` | `None` | Resume from a specific Excel row (None = all). |
| `MAX_ROWS` | `None` | Limit rows in one run (None = all). |
| `CHROMEDRIVER_PATH` | `None` | Use a specific driver (None = auto). |
| `EXPLICIT_WAIT` | `15` | Default wait timeout in seconds. |

---

## Output

- `bulk_add_questions.log` – full run log (per-row success/failure).
- `failure_screenshots/row_<N>_failure.png` – screenshot on any failure.

---

## How it works (high level)

1. Attach to the Chrome tab on port 9222.
2. For each Excel row: fill Question → Options (+ correct radio) → Category →
   Difficulty → Randomize Options → wait for Submit to enable → Submit.
3. After Submit, navigate: *Questions list → Create Individual Questions →
   Single Select* to reach a fresh blank form.
4. On any row error, recover to a fresh form and continue.
5. After the last row, click **Save & Continue** to persist to the server.

---

## Troubleshooting

- **"Excel file not found"** – put `Question_Bank.xlsx` in the folder where you
  opened the terminal, or pass the full path as an argument.
- **Can't connect to Chrome** – make sure Chrome was launched with
  `--remote-debugging-port=9222` and the tab is on the question form.
- **A row failed** – check `bulk_add_questions.log` and the screenshot in
  `failure_screenshots/`. The script keeps going with the next row.
- **Questions didn't persist** – the script clicks "Save & Continue" at the
  end; if that step failed, open the exam's Question step and click it manually.
- **Selectors broke after a platform update** – the CSS/XPath selectors are
  documented in the docstring at the top of `bulk_add_questions.py`; re-verify
  with browser dev tools.

---

## Notes / limitations

- Built and verified against the PULMS "Single Select" question form. Other
  question types (Matching, Multiple Select, etc.) are not handled.
- The script does not launch or close your browser; it only drives the tab.
- Use responsibly and only on exams you own/manage.
