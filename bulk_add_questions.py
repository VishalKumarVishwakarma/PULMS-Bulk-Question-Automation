"""
PULMS Bulk Question Uploader
=============================

Reads questions from an Excel file and automates the "Add Question" form
on an already-open, already-logged-in Chrome tab (attached via remote
debugging on port 9222).

--------------------------------------------------------------------------
BEFORE RUNNING
--------------------------------------------------------------------------
1. Close all Chrome windows, then relaunch Chrome with remote debugging:

       chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\chrome-debug-profile"

   (Use a separate --user-data-dir the first time so it doesn't collide
   with your normal profile lock. Log in manually and navigate to the
   "Create Question" / "Add Question" screen.)

2. Download the chromedriver version that matches your installed Chrome
   version from https://googlechromelabs.github.io/chrome-for-testing/
   and note its path (or put it on PATH).

3. pip install selenium openpyxl pandas

4. Fill in the CONFIG block below (paths, category/sub-category, excel
   column meaning for "Correct Option").

--------------------------------------------------------------------------
ASSUMPTIONS / DEFAULTS BAKED IN (change in CONFIG if wrong)
--------------------------------------------------------------------------
- "Correct Option" column may be either a number 1-4 OR text matching one
  of the Option 1-4 cells. The script auto-detects which.
- Category and Sub Category are NOT per-row Excel columns. They are fixed
  for the whole run (CATEGORY_NAME / SUBCATEGORY_TITLE below), since your
  screenshots showed these look like course/batch groupings rather than
  per-question values. If they DO vary per row, add "Category" and
  "Sub Category" columns to the Excel and the script will prefer those.
- After Submit, the app navigates to a "Questions list" screen (CONFIRMED
  live). The script automatically clicks "Create Individual Questions"
  then "Single Select" to get back to a fresh blank form for the next
  row - no config needed for this.
- Randomize Options toggle and Point if correct / incorrect fields are
  left at whatever default the form already has - not touched.

--------------------------------------------------------------------------
SELECTORS
--------------------------------------------------------------------------
Confirmed directly against the live DOM (elearning.paruluniversity.ac.in
/exams/edit/...) via an attached Selenium session on 2026-07-22:
- Question/Option fields: Quill.js editors, ".ql-editor"
- "Is correct" radios: label[formcontrolname='isCorrect'] input.ant-radio-input
- "Add Option" button: ion-button with exact text "Add Option"
- Category / Sub Category: separate nz-select (ng-zorro) fields, each
  with its own "Add Category" / "Add Sub Category" modal
  (formcontrolname='name' for Category name; a plain text input under
  the "Sub Category Title" label for Sub Category)
- Difficulty radios: #Easy / #Medium / #Hard
- Submit: ion-button with exact text "Submit", disabled via disabled/
  aria-disabled attributes while the form is invalid
If the app is updated and markup changes, re-verify with dev tools.
"""

import logging
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
)
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# ==========================================================================
# CONFIG - edit these before running
# ==========================================================================

# Name of the Excel question bank the script looks for.
EXCEL_FILENAME = "Question_Bank.xlsx"

# Directory this script file lives in (used as a fallback location).
SCRIPT_DIR = Path(__file__).resolve().parent


def _resolve_excel_path():
    """Find the question bank so the script can be run from ANY folder via
    cmd. Resolution order:
      1. A path passed as the first command-line argument
         (e.g. `python bulk_add_questions.py "D:\\my questions.xlsx"`).
      2. "Question_Bank.xlsx" in the current working directory (the folder
         where you opened cmd). <- normal case.
      3. "Question_Bank.xlsx" sitting next to this script file.
    If none exist, returns the cwd path so the later existence check can
    show a clear error.
    """
    if len(sys.argv) > 1 and sys.argv[1].strip():
        return Path(sys.argv[1]).expanduser().resolve()
    cwd_candidate = Path.cwd() / EXCEL_FILENAME
    if cwd_candidate.exists():
        return cwd_candidate.resolve()
    script_candidate = SCRIPT_DIR / EXCEL_FILENAME
    if script_candidate.exists():
        return script_candidate.resolve()
    return cwd_candidate.resolve()


EXCEL_PATH = str(_resolve_excel_path())
# Log file and failure screenshots are written next to the Excel file, so
# outputs land wherever you run from.
OUTPUT_DIR = Path(EXCEL_PATH).parent

# No chromedriver.exe was found on this machine, so we rely on Selenium
# Manager (built into Selenium 4.6+) to auto-download the matching driver.
# If you have a specific chromedriver.exe you want to use instead, set its
# path here and it will be used; otherwise leave as None.
CHROMEDRIVER_PATH = None
DEBUGGER_ADDRESS = "127.0.0.1:9222"

# Category / Sub Category used for every row unless the Excel has its own
# "Category" / "Sub Category" columns.
CATEGORY_NAME = "Weekly_exam"
SUBCATEGORY_TITLE = None  # None/empty -> Sub Category field is left untouched

# Turn the "Randomize Options" toggle ON for every question. Set to False
# to leave it at its default (off).
RANDOMIZE_OPTIONS = True

# CONFIRMED live behavior: after Submit, the app does NOT reset the form
# in place. It navigates to a "Questions list" screen showing all
# questions added so far, with a "Create Individual Questions" button.
# Clicking that opens a "Select Question Type" screen, and clicking
# "Single Select" there opens a fresh blank single-choice question form.
# This 2-step navigation is handled by navigate_to_fresh_form().

# Start processing from this Excel row number (matches the "row N" shown
# in the log; the header is row 1, so the first data row is row 2).
# Set to None (default) to process EVERY question in the Excel file.
# (Only set a value if you ever need to resume a partial run.)
START_ROW = None

# Process at most this many rows in one run (after applying START_ROW).
# Set to None (default) to process ALL rows - i.e. the whole question
# bank. (Only set a number for a small test run.)
MAX_ROWS = None

EXPLICIT_WAIT = 15  # seconds, default explicit wait timeout
SCREENSHOT_DIR = OUTPUT_DIR / "failure_screenshots"

# ==========================================================================
# LOGGING
# ==========================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(
            OUTPUT_DIR / "bulk_add_questions.log", encoding="utf-8"
        ),
    ],
)
log = logging.getLogger("pulms")


# ==========================================================================
# DATA MODEL
# ==========================================================================

@dataclass
class QuestionRow:
    row_number: int  # 1-based, matches Excel row (for logging)
    question: str
    options: list  # list of non-empty option strings, in order
    correct_index: int  # 0-based index into `options`
    difficulty: str  # "Easy" | "Medium" | "Hard"
    category: str
    subcategory: str


def load_rows(excel_path: str, default_category: str, default_subcategory: str):
    df = pd.read_excel(excel_path, dtype=str).fillna("")
    required_cols = {
        "Question",
        "Option 1",
        "Option 2",
        "Option 3",
        "Option 4",
        "Correct Option",
        "Difficulty Level",
    }
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"Excel is missing required columns: {missing}")

    has_category_col = "Category" in df.columns
    has_subcategory_col = "Sub Category" in df.columns

    rows = []
    for i, r in df.iterrows():
        row_number = i + 2  # +1 for 0-index, +1 for header row
        question = str(r["Question"]).strip()
        if not question:
            log.warning("Row %s has an empty Question, skipping.", row_number)
            continue

        raw_options = [
            str(r["Option 1"]).strip(),
            str(r["Option 2"]).strip(),
            str(r["Option 3"]).strip(),
            str(r["Option 4"]).strip(),
        ]
        options = [o for o in raw_options if o]
        if len(options) < 2:
            log.warning(
                "Row %s has fewer than 2 non-empty options, skipping.", row_number
            )
            continue

        correct_raw = str(r["Correct Option"]).strip()
        correct_index = resolve_correct_index(correct_raw, raw_options, row_number)
        if correct_index is None:
            log.warning(
                "Row %s: could not resolve Correct Option %r against options, skipping.",
                row_number,
                correct_raw,
            )
            continue

        difficulty = str(r["Difficulty Level"]).strip().capitalize()
        if difficulty not in ("Easy", "Medium", "Hard"):
            log.warning(
                "Row %s has unrecognized Difficulty Level %r, defaulting to Easy.",
                row_number,
                difficulty,
            )
            difficulty = "Easy"

        category = (
            str(r["Category"]).strip()
            if has_category_col and str(r["Category"]).strip()
            else default_category
        )
        subcategory = (
            str(r["Sub Category"]).strip()
            if has_subcategory_col and str(r["Sub Category"]).strip()
            else default_subcategory
        )

        rows.append(
            QuestionRow(
                row_number=row_number,
                question=question,
                options=options,
                correct_index=correct_index,
                difficulty=difficulty,
                category=category,
                subcategory=subcategory,
            )
        )
    return rows


_OPTION_N_RE = re.compile(r"^\s*option\s*([1-4])\s*$", re.IGNORECASE)


def resolve_correct_index(correct_raw: str, raw_options: list, row_number: int):
    """Correct Option can be:
      - a bare number: '1'-'4' (1-based column number)
      - the literal option text itself
      - the phrase 'Option N' (e.g. 'Option 2') - confirmed format used in
        BDA_Unit1_Unit2_MCQ_QuestionBank.xlsx
    raw_options is the ORIGINAL 4-slot list (with blanks); correct_index
    returned is relative to the compacted non-empty list."""
    non_empty_positions = [idx for idx, o in enumerate(raw_options) if o]

    if correct_raw.isdigit():
        col_index = int(correct_raw) - 1  # 0-based within the original 4 slots
        if col_index in non_empty_positions:
            return non_empty_positions.index(col_index)
        return None

    m = _OPTION_N_RE.match(correct_raw)
    if m:
        col_index = int(m.group(1)) - 1
        if col_index in non_empty_positions:
            return non_empty_positions.index(col_index)
        return None

    # fall back to text match (case-insensitive, exact)
    for pos, opt in enumerate(raw_options):
        if opt and opt.strip().lower() == correct_raw.strip().lower():
            if pos in non_empty_positions:
                return non_empty_positions.index(pos)
    return None


# ==========================================================================
# DRIVER SETUP
# ==========================================================================

def get_driver():
    options = Options()
    options.debugger_address = DEBUGGER_ADDRESS
    if CHROMEDRIVER_PATH:
        service = webdriver.chrome.service.Service(executable_path=CHROMEDRIVER_PATH)
        driver = webdriver.Chrome(service=service, options=options)
    else:
        # Selenium Manager auto-resolves a matching chromedriver.
        driver = webdriver.Chrome(options=options)
    return driver


# ==========================================================================
# FORM AUTOMATION HELPERS
# ==========================================================================

def wait_visible(driver, by, value, timeout=EXPLICIT_WAIT):
    """Waits for presence first, scrolls into view (see wait_clickable for
    why), then waits for Selenium's visibility check to pass."""
    element = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, value))
    )
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
    return WebDriverWait(driver, timeout).until(
        EC.visibility_of_element_located((by, value))
    )


def wait_clickable(driver, by, value, timeout=EXPLICIT_WAIT):
    """Waits for the element to be present, scrolls it into view (Selenium's
    is_displayed()/element_to_be_clickable() report elements below the
    viewport as not visible, which caused indefinite hangs on this page's
    long single-page form), then waits for it to be clickable."""
    element = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, value))
    )
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
    return WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((by, value))
    )


def wait_all(driver, by, value, timeout=EXPLICIT_WAIT):
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_all_elements_located((by, value))
    )


def safe_click(driver, element):
    """Scroll the element into view (required on this form - elements
    below the viewport are treated as not displayed/not clickable by
    Selenium) then click, falling back to a JS click if a real click is
    intercepted (common with Angular overlays/animations)."""
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
    try:
        element.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", element)


def get_rich_text_editors(driver):
    """Returns all rich-text editable divs on the form, in DOM order.
    Index 0 is the Question editor, index 1..N are the Option editors.

    CONFIRMED against the live page: editors use Quill.js (".ql-editor").
    There is also a hidden ".ql-clipboard" contenteditable per editor
    instance, which is why we do NOT fall back to a bare
    "[contenteditable='true']" selector (that would double-count).
    """
    return driver.find_elements(By.CSS_SELECTOR, ".ql-editor")


def clear_and_type_rich_text(driver, editor_element, text):
    safe_click(driver, editor_element)
    # select all existing content inside this editable div and delete it
    editor_element.send_keys(Keys.CONTROL, "a")
    editor_element.send_keys(Keys.DELETE)
    editor_element.send_keys(text)


def fill_question(driver, wait, question_text):
    editors = get_rich_text_editors(driver)
    if not editors:
        raise NoSuchElementException("No rich-text editors found for Question field.")
    clear_and_type_rich_text(driver, editors[0], question_text)


def get_add_option_button(driver):
    # CONFIRMED: unique ion-button with exact text "Add Option".
    candidates = driver.find_elements(
        By.XPATH, "//ion-button[normalize-space(.)='Add Option']"
    )
    return candidates[0] if candidates else None


def get_option_radio_inputs(driver):
    # CONFIRMED: "Is correct" radios use input.ant-radio-input, scoped
    # inside label[formcontrolname='isCorrect']. Difficulty radios are
    # plain <input type="radio" id="Easy/Medium/Hard"> without this class,
    # so there's no collision - this selector is safe as-is.
    return driver.find_elements(
        By.CSS_SELECTOR, "label[formcontrolname='isCorrect'] input.ant-radio-input"
    )


def fill_options(driver, wait, options: list, correct_index: int):
    needed = len(options)

    # Ensure enough option blocks exist. Assume the form starts with 1
    # option block (Option 1) already present.
    for _ in range(needed - 1):
        btn = get_add_option_button(driver)
        if btn is None:
            raise NoSuchElementException("Could not find 'Add Option' button.")
        safe_click(driver, btn)
        time.sleep(0.4)  # brief settle for Angular re-render before next click

    # Re-fetch editors after all option blocks exist. Editor[0] = Question,
    # editors[1:1+needed] = options in order.
    editors = get_rich_text_editors(driver)
    option_editors = editors[1 : 1 + needed]
    if len(option_editors) < needed:
        raise NoSuchElementException(
            f"Expected {needed} option editors, found {len(option_editors)}."
        )

    for text, editor in zip(options, option_editors):
        clear_and_type_rich_text(driver, editor, text)

    radios = get_option_radio_inputs(driver)
    if len(radios) < needed:
        raise NoSuchElementException(
            f"Expected {needed} 'Is correct' radios, found {len(radios)}."
        )
    safe_click(driver, radios[correct_index])


# CONFIRMED xpaths: the Category field is a "p" label containing exactly
# "Category" (not "Sub Category"), and Sub Category is scoped separately.
# Both wrap an ng-zorro nz-select with the standard search input class.
_CATEGORY_INPUT_XPATH = (
    "//p[contains(., 'Category') and not(contains(.,'Sub'))]"
    "/ancestor::div[contains(@class,'flex-column')][1]"
    "//input[contains(@class,'ant-select-selection-search-input')]"
)
_SUBCATEGORY_INPUT_XPATH = (
    "//p[contains(., 'Sub Category')]"
    "/ancestor::div[contains(@class,'flex-column')][1]"
    "//input[contains(@class,'ant-select-selection-search-input')]"
)


def _select_dropdown_option_or_none(driver, option_text, timeout=10):
    """Selects an ant-select dropdown option whose exact text (title
    attribute) equals option_text. Returns True if found+clicked, else
    False.

    CONFIRMED bugs this handles (found by live testing 2026-07-22):
    - The dropdown uses server-side search (nzserversearch) with a
      transient "Loading..." item; results can take a couple seconds.
    - Options live inside a cdk-virtual-scroll viewport, so many are NOT
      rendered/visible even when present in the data. Selenium's
      visibility check therefore fails on them, and a real .click()
      raises "element not interactable". So we wait on PRESENCE (not
      visibility), match on the reliable @title attribute, and click via
      JavaScript (which works where a native click does not).
    """
    safe_text = option_text.replace('"', '\\"')
    option_xpath = f"//nz-option-item[@title=\"{safe_text}\"]"
    try:
        option = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, option_xpath))
        )
    except TimeoutException:
        return False
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", option)
    driver.execute_script("arguments[0].click();", option)
    time.sleep(0.4)
    return True


def _type_search(driver, search_input, text):
    """Types into an ng-zorro server-search select input one character at
    a time. CONFIRMED bug: sending the whole string at once
    (send_keys(text)) does NOT trigger Angular's (valueChanges) search,
    so the list stays unfiltered; typing char-by-char with small delays
    reliably fires the search."""
    for ch in text:
        search_input.send_keys(ch)
        time.sleep(0.08)


def _clear_search_input(driver, search_input):
    """Clears an ng-zorro search-select input. Plain .clear() does not
    reliably work on these Angular-bound inputs, and if the input isn't
    cleared before typing, repeated calls accumulate text and the
    dropdown returns no matches at all."""
    search_input.send_keys(Keys.CONTROL, "a")
    search_input.send_keys(Keys.DELETE)
    time.sleep(0.2)


def _close_open_modal(driver):
    icons = driver.find_elements(By.CSS_SELECTOR, ".ant-modal-header .icon-close")
    if icons:
        driver.execute_script("arguments[0].closest('ion-button').click();", icons[0])
        time.sleep(0.4)


def set_category(driver, wait, category_name):
    search_input = wait_clickable(driver, By.XPATH, _CATEGORY_INPUT_XPATH)
    safe_click(driver, search_input)
    _clear_search_input(driver, search_input)
    _type_search(driver, search_input, category_name)

    if _select_dropdown_option_or_none(driver, category_name):
        return

    # Not found - create it via the "Add Category" modal.
    log.info("Category %r not found in dropdown, creating it via modal.", category_name)
    search_input.send_keys(Keys.ESCAPE)  # close the empty dropdown first
    add_btn = wait_clickable(
        driver, By.XPATH, "//ion-button[normalize-space(.)='Add Category']"
    )
    safe_click(driver, add_btn)

    name_input = wait_visible(driver, By.CSS_SELECTOR, "input[formcontrolname='name']")
    name_input.clear()
    name_input.send_keys(category_name)

    add_button = wait_clickable(
        driver, By.XPATH, "//div[contains(@class,'ant-modal-footer')]//ion-button[normalize-space(.)='Add']"
    )
    safe_click(driver, add_button)
    time.sleep(0.6)
    _close_open_modal(driver)  # in case modal doesn't auto-close after Add

    # After creating, re-select it from the (now updated) dropdown.
    safe_click(driver, search_input)
    _clear_search_input(driver, search_input)
    _type_search(driver, search_input, category_name)
    if not _select_dropdown_option_or_none(driver, category_name):
        raise NoSuchElementException(
            f"Category '{category_name}' still not selectable after creating it."
        )


def set_subcategory(driver, wait, subcategory_title, category_name):
    """Sub Category field. Always present on this form per confirmed DOM
    structure, but we still guard with a try in case it's absent on some
    exam types."""
    try:
        sub_search_input = wait_clickable(driver, By.XPATH, _SUBCATEGORY_INPUT_XPATH, timeout=4)
    except TimeoutException:
        log.info("No Sub Category field found on the form, skipping.")
        return

    safe_click(driver, sub_search_input)
    _clear_search_input(driver, sub_search_input)
    _type_search(driver, sub_search_input, subcategory_title)

    if _select_dropdown_option_or_none(driver, subcategory_title):
        return

    log.info("Sub Category %r not found, creating it via modal.", subcategory_title)
    sub_search_input.send_keys(Keys.ESCAPE)
    add_btn = wait_clickable(
        driver, By.XPATH, "//ion-button[normalize-space(.)='Add Sub Category']"
    )
    safe_click(driver, add_btn)

    # Modal fields: "Sub Category Title" (plain text input) and an
    # optional "Category" nz-select inside the modal itself.
    title_input = wait_visible(
        driver,
        By.XPATH,
        "//p[contains(., 'Sub Category Title')]/ancestor::div[contains(@class,'flex-column')][1]//input[@type='text']",
    )
    title_input.clear()
    title_input.send_keys(subcategory_title)

    # Best-effort: set the Category dropdown inside the modal to match.
    try:
        modal_cat_input = driver.find_element(
            By.XPATH,
            "//p[contains(., 'Category')]/ancestor::div[contains(@class,'flex-column')][1]//input[contains(@class,'ant-select-selection-search-input')]",
        )
        safe_click(driver, modal_cat_input)
        _clear_search_input(driver, modal_cat_input)
        _type_search(driver, modal_cat_input, category_name)
        _select_dropdown_option_or_none(driver, category_name, timeout=10)
    except NoSuchElementException:
        pass

    add_button = wait_clickable(
        driver, By.XPATH, "//div[contains(@class,'ant-modal-footer')]//ion-button[normalize-space(.)='Add']"
    )
    safe_click(driver, add_button)
    time.sleep(0.6)
    _close_open_modal(driver)

    safe_click(driver, sub_search_input)
    _clear_search_input(driver, sub_search_input)
    _type_search(driver, sub_search_input, subcategory_title)
    if not _select_dropdown_option_or_none(driver, subcategory_title):
        raise NoSuchElementException(
            f"Sub Category '{subcategory_title}' still not selectable after creating it."
        )


def set_randomize_options(driver, wait, turn_on: bool):
    """Turns the "Randomize Options" toggle on/off.

    CONFIRMED live structure: an <ion-toggle> inside
    ed-toggle[formcontrolname='randomizeOptions']; its state is exposed
    via aria-checked ("true"/"false"). We only click when the current
    state differs from the desired one (clicking is a flip)."""
    try:
        toggle = wait_clickable(
            driver,
            By.CSS_SELECTOR,
            "ed-toggle[formcontrolname='randomizeOptions'] ion-toggle",
            timeout=6,
        )
    except TimeoutException:
        log.warning("Randomize Options toggle not found; skipping.")
        return

    current_on = (toggle.get_attribute("aria-checked") == "true")
    if current_on == turn_on:
        return  # already in the desired state
    safe_click(driver, toggle)
    time.sleep(0.3)
    # verify
    new_state = driver.find_element(
        By.CSS_SELECTOR, "ed-toggle[formcontrolname='randomizeOptions'] ion-toggle"
    ).get_attribute("aria-checked") == "true"
    if new_state != turn_on:
        log.warning("Randomize Options toggle did not reach the desired state.")


def set_difficulty(driver, wait, difficulty: str):
    # Matches your screenshot: radio ids #Easy / #Medium / #Hard.
    radio = wait_clickable(driver, By.CSS_SELECTOR, f"#{difficulty}")
    safe_click(driver, radio)


def submit_form(driver, wait):
    submit_btn = WebDriverWait(driver, EXPLICIT_WAIT).until(
        lambda d: _get_enabled_submit_button(d)
    )
    safe_click(driver, submit_btn)


def _get_enabled_submit_button(driver):
    # CONFIRMED: Submit is an ion-button with exact text "Submit". When
    # disabled it carries both disabled="" and aria-disabled="true".
    buttons = driver.find_elements(
        By.XPATH, "//ion-button[normalize-space(.)='Submit']"
    )
    if not buttons:
        return False
    btn = buttons[0]
    disabled = btn.get_attribute("disabled")
    aria_disabled = btn.get_attribute("aria-disabled")
    if disabled is None and aria_disabled in (None, "false"):
        return btn
    return False


def navigate_to_fresh_form(driver, wait):
    """CONFIRMED live flow after Submit: the app shows a "Questions list"
    screen. To add the next question we must:
      1. Click "Create Individual Questions"
      2. Click "Single Select" on the "Select Question Type" screen
    This lands back on a fresh blank single-choice question form (2
    empty .ql-editor instances: Question + Option 1)."""
    create_btn = wait_clickable(
        driver, By.XPATH, "//ion-button[contains(., 'Create Individual Questions')]"
    )
    safe_click(driver, create_btn)

    # The "Single Select" tile is an <h6>, not a real button, so a native
    # click can be flaky - locate it (presence) and JS-click.
    single_select = WebDriverWait(driver, EXPLICIT_WAIT).until(
        EC.presence_of_element_located(
            (By.XPATH, "//h6[normalize-space(text())='Single Select']")
        )
    )
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", single_select)
    driver.execute_script("arguments[0].click();", single_select)

    try:
        WebDriverWait(driver, EXPLICIT_WAIT).until(_blank_form_ready)
    except TimeoutException:
        log.warning(
            "Could not confirm a fresh blank form loaded after navigation; continuing anyway."
        )


# ==========================================================================
# MAIN PROCESSING
# ==========================================================================

def process_row(driver, wait, row: QuestionRow) -> bool:
    fill_question(driver, wait, row.question)
    fill_options(driver, wait, row.options, row.correct_index)
    set_category(driver, wait, row.category)
    if row.subcategory:
        set_subcategory(driver, wait, row.subcategory, row.category)
    else:
        log.info("Sub Category left untouched (SUBCATEGORY_TITLE not set).")
    set_difficulty(driver, wait, row.difficulty)
    set_randomize_options(driver, wait, RANDOMIZE_OPTIONS)
    submit_form(driver, wait)
    navigate_to_fresh_form(driver, wait)
    return True


def reach_questions_list(driver, wait, max_backs=4):
    """Navigate (via the form's "Back" button) to the "Questions list"
    screen, identified by the presence of "Create Individual Questions".
    Returns True if reached."""
    for _ in range(max_backs):
        if driver.find_elements(
            By.XPATH, "//ion-button[contains(., 'Create Individual Questions')]"
        ):
            return True
        back_btns = driver.find_elements(
            By.XPATH, "//ion-button[normalize-space(.)='Back']"
        )
        if not back_btns:
            break
        safe_click(driver, back_btns[0])
        time.sleep(1.5)
    return bool(
        driver.find_elements(
            By.XPATH, "//ion-button[contains(., 'Create Individual Questions')]"
        )
    )


def persist_questions(driver, wait):
    """CONFIRMED essential: questions added via the form's Submit are only
    held in the wizard's session list. They are persisted to the server
    ONLY when "Save & Continue" is clicked on the Questions-list screen
    (verified: after Save & Continue the exam list showed the real
    question count; without it, a reload lost the questions).

    This navigates to the Questions list and clicks "Save & Continue".
    Returns True on success.
    """
    if not reach_questions_list(driver, wait):
        log.error("persist_questions: could not reach the Questions list to save.")
        return False
    save_btn = driver.find_elements(
        By.XPATH, "//ion-button[normalize-space(.)='Save & Continue']"
    )
    if not save_btn:
        log.error("persist_questions: 'Save & Continue' button not found.")
        return False
    safe_click(driver, save_btn[0])
    time.sleep(2.5)
    log.info("Clicked 'Save & Continue' to persist questions to the server.")
    return True


def _blank_form_ready(driver):
    """True if a fresh, empty single-choice question form is showing."""
    editors = get_rich_text_editors(driver)
    if len(editors) < 2:
        return False
    try:
        return driver.execute_script("return arguments[0].innerText.trim() === ''", editors[0])
    except Exception:
        return False


def _dismiss_overlays(driver):
    """Close any open dropdown (ESC) or modal so navigation buttons are
    reachable."""
    from selenium.webdriver.common.action_chains import ActionChains

    try:
        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        time.sleep(0.3)
    except Exception:  # noqa: BLE001
        pass
    _close_open_modal(driver)


def recover_to_fresh_form(driver, wait):
    """After a failed row the browser can be left in many states: a
    dropdown open, a modal open, a half-filled form, the "Select Question
    Type" screen, or the questions list. This walks the SPA back to a
    guaranteed-fresh blank question form WITHOUT reloading the page
    (reload resets the whole exam wizard to step 1, which is fragile).

    The original run cascaded every remaining row into failure because
    there was no recovery at all; this fixes that.
    """
    _dismiss_overlays(driver)

    for attempt in range(6):
        # Already on a "Select Question Type" screen? Click Single Select.
        if driver.find_elements(By.XPATH, "//h6[normalize-space(text())='Single Select']"):
            try:
                navigate_single_select(driver)
                if _blank_form_ready(driver):
                    return
            except Exception:  # noqa: BLE001
                pass

        # On the questions list? Go create a fresh individual question.
        if driver.find_elements(
            By.XPATH, "//ion-button[contains(., 'Create Individual Questions')]"
        ):
            try:
                navigate_to_fresh_form(driver, wait)
                if _blank_form_ready(driver):
                    return
            except Exception:  # noqa: BLE001
                pass

        # On a (dirty) form? Use its "Back" button to reach the list.
        back_btns = driver.find_elements(
            By.XPATH, "//ion-button[normalize-space(.)='Back']"
        )
        if back_btns:
            safe_click(driver, back_btns[0])
            time.sleep(1.2)
            _dismiss_overlays(driver)
            continue

        # Otherwise try clicking the "Question" wizard step to reach list.
        steps = driver.find_elements(
            By.XPATH,
            "//div[contains(@class,'ant-steps-item')][.//text()[contains(.,'Question')]]",
        )
        if steps:
            safe_click(driver, steps[0])
            time.sleep(1.5)
            continue

        time.sleep(0.5)

    log.error("Recovery could not reach a fresh blank form after several attempts.")


def navigate_single_select(driver):
    """Clicks the "Single Select" tile and waits for the blank form."""
    single_select = WebDriverWait(driver, EXPLICIT_WAIT).until(
        EC.presence_of_element_located(
            (By.XPATH, "//h6[normalize-space(text())='Single Select']")
        )
    )
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", single_select)
    driver.execute_script("arguments[0].click();", single_select)
    WebDriverWait(driver, EXPLICIT_WAIT).until(_blank_form_ready)


def _try_recover(driver, wait, row_number):
    """Wraps recover_to_fresh_form so a recovery failure never aborts the
    whole batch."""
    try:
        recover_to_fresh_form(driver, wait)
    except Exception:  # noqa: BLE001
        log.exception("Recovery after row %s failed; will still attempt next row.", row_number)


def save_failure_screenshot(driver, row_number: int):
    SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)
    path = SCREENSHOT_DIR / f"row_{row_number}_failure.png"
    try:
        driver.save_screenshot(str(path))
        log.info("Saved failure screenshot to %s", path)
    except Exception as e:
        log.error("Could not save screenshot: %s", e)


def main():
    log.info("Using question bank: %s", EXCEL_PATH)
    if not Path(EXCEL_PATH).exists():
        log.error("Excel file not found: %s", EXCEL_PATH)
        log.error(
            "Put '%s' in the folder where you open cmd (or pass its full path "
            "as an argument: python bulk_add_questions.py \"C:\\path\\to\\file.xlsx\").",
            EXCEL_FILENAME,
        )
        sys.exit(1)

    rows = load_rows(EXCEL_PATH, CATEGORY_NAME, SUBCATEGORY_TITLE)
    log.info("Loaded %d valid question rows from Excel.", len(rows))
    if not rows:
        log.error("No valid rows to process. Exiting.")
        sys.exit(1)

    driver = get_driver()
    wait = WebDriverWait(driver, EXPLICIT_WAIT)

    # Make the run robust to the exact starting screen: if the page is on
    # the questions list or a "Select Question Type" screen (instead of a
    # blank question form), navigate to a fresh blank form first so the
    # very first row does not fail.
    if not _blank_form_ready(driver):
        log.info("Not on a blank question form at start; navigating to one.")
        try:
            recover_to_fresh_form(driver, wait)
        except Exception:  # noqa: BLE001
            log.exception("Could not reach a blank form at start; will still try.")

    success_count = 0
    failure_count = 0

    if START_ROW is not None:
        rows = [r for r in rows if r.row_number >= START_ROW]
        log.info("START_ROW=%s -> processing %d rows from Excel row %s onward.",
                 START_ROW, len(rows), START_ROW)
    if MAX_ROWS is not None:
        rows = rows[:MAX_ROWS]
        log.info("MAX_ROWS=%s -> limiting this run to %d rows.", MAX_ROWS, len(rows))

    for row in rows:
        preview = row.question[:60] + ("..." if len(row.question) > 60 else "")
        log.info("Processing row %s: %r", row.row_number, preview)
        try:
            process_row(driver, wait, row)
            success_count += 1
            log.info("Row %s SUCCESS: %r", row.row_number, preview)
        except (
            TimeoutException,
            NoSuchElementException,
            StaleElementReferenceException,
            ElementClickInterceptedException,
        ) as e:
            failure_count += 1
            log.error("Row %s FAILED: %r -> %s: %s", row.row_number, preview, type(e).__name__, e)
            save_failure_screenshot(driver, row.row_number)
            _try_recover(driver, wait, row.row_number)
        except Exception as e:  # noqa: BLE001 - log unexpected errors too, keep going
            failure_count += 1
            log.exception("Row %s FAILED with unexpected error: %r", row.row_number, preview)
            save_failure_screenshot(driver, row.row_number)
            _try_recover(driver, wait, row.row_number)

    # Persist everything to the server. Questions added via Submit only
    # live in the wizard session until "Save & Continue" is clicked; the
    # script never reloads mid-run, so all successful questions are still
    # in the session here and this one save commits them.
    if success_count > 0:
        try:
            persist_questions(driver, wait)
        except Exception:  # noqa: BLE001
            log.exception("Failed to click 'Save & Continue'. Questions may NOT be "
                          "saved - open the exam, go to the Questions step, and click "
                          "'Save & Continue' manually.")

    log.info("Done. Success: %d, Failed: %d, Total: %d", success_count, failure_count, len(rows))
    # NOTE: We do not quit the driver since it's attached to a Chrome
    # instance the user is still using interactively.


if __name__ == "__main__":
    main()
