# VistA Radiology Quick Order Automation

A Python script that automates the creation of radiology "quick orders" in VistA, using screen-reading and simulated keystrokes against an already-open Reflection Desktop Pro / EXTRA! terminal session. Built to support bulk quick order creation as part of a HICOP requiring a large number of new imaging orders across multiple new menus.

This does **not** connect to VistA over the network, bypass any authentication, or modify VistA itself. It automates the same manual keystrokes a person would type into an already-logged-in terminal session, driven from a simple two-column spreadsheet.

---


## Table of Contents

- [How It Works](#how-it-works)
- [Setup Guide](#setup-guide)
- [Running the Script](#running-the-script)
- [Configuration Reference](#configuration-reference)
- [Troubleshooting](#troubleshooting)
- [Process Notes & Lessons Learned](#process-notes--lessons-learned)

---

## How It Works

**This script cannot create an order for a procedure that doesn't exist in VistA yet.** Radiology must add the procedure name into VistA (in whichever environment you're running against) before the script can select it — this is a hard prerequisite, not a timing nuance. Trying to run an order before radiology has added its procedure will result in VistA rejecting the procedure name outright, every time, regardless of how the script is configured.

For each row in your spreadsheet, the script:

1. Waits for VistA's `Select QUICK ORDER NAME:` prompt, then types a full order name built from a configurable prefix + your order name + a configurable suffix.
2. Walks through the full quick order creation sequence — confirming the new order, selecting "Imaging" as the order type, entering the display text, selecting the imaging type, and entering the procedure name — waiting for each specific VistA prompt to appear before typing anything.
3. Leaves fields like Verify Order, Description, Procedure Modifier, and Reason for Study blank, matching local convention.
4. Pastes a fixed Clinical History template from the clipboard (you copy this once before running).
5. Leaves all remaining fields (Category, Pre-op, Date Desired, Mode of Transport, Isolation, Urgency, Submit To, Place/Auto-accept) at their defaults.
6. Confirms the order was placed and moves to the next row.

The script never assumes a fixed timing — every step waits for actual confirmation from the VistA screen before proceeding, and every typed field is read back and verified to make sure it landed correctly. If anything doesn't match what's expected, the script pauses, beeps, and hands control back to you so you can resolve it manually in VistA, then resume by pressing **ESC**.

---

## Setup Guide

### Before You Start

- [Visual Studio Code](https://code.visualstudio.com/) with the Python extension installed, and the ability to create a Python virtual environment
- Reflection Desktop Pro (or EXTRA!) installed and connected to your **test** VistA account
- Microsoft Excel or similar, to build your order spreadsheet
- **Confirmation from your radiology department that each procedure on your spreadsheet has already been added into VistA.** This script only automates the keystrokes of creating a quick order — it cannot create the underlying procedure itself. Any order attempted before radiology has added its procedure will fail, in test or production alike. Coordinate with radiology before building your spreadsheet so you're only including procedures that are actually ready.
- About 20–30 minutes for first-time setup

**Always test in your site's test account first.** Never run an unfamiliar or freshly modified version of this script against production.

**Important — test account data may be out of date.** Your VistA test account is typically a snapshot/backup of production taken at some point in the past — it is **not** automatically kept in sync with new procedure names as radiology adds them to production. If radiology has added new procedure names to production *after* your test account's last backup/refresh, those procedure names will not exist in test yet, and the script will fail there — not because of a bug, but because VistA in test genuinely has no matching procedure to select. This can look identical to a real script problem (a rejected procedure name, an unexpected prompt) even though the script itself is working correctly.

If you hit unexplained procedure rejection errors in test, check with radiology or your VistA team on whether the new procedure names actually exist in the test environment yet before assuming the script is broken.

Because of this gap, testing may need to happen directly in **production** in practice. If so, set `ORDER_PREFIX` in the script to a clearly distinct, obviously-fake value (e.g. `"ZZTEST "`) instead of your real production prefix (e.g. `"RAZN "`). This keeps test orders created during development easy to identify and remove, and prevents them from being mistaken for real, intended quick orders. Only change `ORDER_PREFIX` to your real production value once you're confident the script is behaving correctly.

**Note on Git:** If your site's IT permissions don't allow installing Git, that's fine — you don't need it. You can copy the script's code directly from this GitHub page instead of cloning the repo. Instructions below use that approach.

### 1. Install Python

1. Go to [python.org/downloads](https://www.python.org/downloads/) and download the Windows installer.
2. Run it. On the first screen, **check "Add python.exe to PATH"** before clicking Install. This step is easy to miss and breaks everything downstream if skipped.
3. Click Install Now, then Close.
4. Confirm it worked: open VS Code, open a new terminal (**Terminal → New Terminal** from the top menu), and run:
   ```
   python --version
   ```
   You should see something like `Python 3.12.x`.

If your site doesn't allow installing new software at all (including Python), check with your local IT team first — this script cannot run without a working Python installation.

### 2. Get the script's code

Since installing Git may not be an option:

1. Go to this repository's page on GitHub.
2. Click on `radiologyOrderBuilder.py` in the file list.
3. Click the **"Raw"** button near the top of the file view — this shows the plain, unformatted code.
4. Select all the text (Ctrl+A) and copy it (Ctrl+C).
5. In VS Code, create a new file (**File → New Text File**), paste the code in, and save it as `radiologyOrderBuilder.py` in a project folder of your choice (e.g. `Desktop\RadOrderAutomation`).

Repeat this same copy/paste process for any other files in the repo you want locally (this README, for reference).

### 3. Create a virtual environment in VS Code

1. In VS Code, open your project folder (**File → Open Folder**, select `RadOrderAutomation`).
2. Open a terminal inside VS Code (**Terminal → New Terminal**).
3. Create the virtual environment:
   ```
   python -m venv .venv
   ```
4. VS Code will likely pop up a notification asking if you want to use this new environment as your workspace's Python interpreter — click **Yes**. If it doesn't ask, you can select it manually: press `Ctrl+Shift+P`, type "Python: Select Interpreter," and choose the one showing `.venv`.
5. Open a new terminal (close the old one and open a fresh one via **Terminal → New Terminal**) so it activates the environment automatically. You'll know it's active when you see `(.venv)` at the start of the terminal prompt.

### 4. Install required packages

In the VS Code terminal, with `(.venv)` showing:

```
pip install pywin32 pywinauto pyperclip pandas openpyxl
```

### 5. Build your order spreadsheet

Create an Excel file with exactly these two column headers in row 1:

| order_name | imaging_type |
|---|---|

Each row is one order. `order_name` is used for both the display name and the procedure name (they're assumed identical). `imaging_type` must exactly match one of:

- `CT SCAN`
- `MAGNETIC RESONANCE IMAGING` (or `MRI`)
- `ULTRASOUND`
- `MAMMOGRAPHY`
- `NUCLEAR MEDICINE` (or `NUC`)
- `GENERAL RADIOLOGY`
- `ANGIO/NEURO/INTERVENTIONAL`
- `VASCULAR LAB`

Save as `.xlsx`.

If you know Vista, less is more often times. This script hard codes these options as acceptable values, but in Vista, you can enter 'C' at the Imaging Type step, and it will give
you 'CT SCAN'. For simplicity, I didn't hard code all the synonyms or abbreviations that you can technically use in Vista. 

### 6. Configure the script

Open `radiologyOrderBuilder.py` in VS Code and edit the values in the `CONFIG` section near the top — see [Configuration Reference](#configuration-reference) below for what each setting does.

---

## Running the Script

1. Open your terminal emulator and log into your **test** VistA account.
2. Navigate through VistA's menus to the **Quick Order entry menu**, until you see the `Select QUICK ORDER NAME:` prompt on screen. The script assumes you're already sitting at this exact prompt when it starts — it does not navigate VistA's menus for you.
3. Copy your Clinical History template text to your clipboard (Ctrl+C) — do this fresh before every run.
4. In VS Code's terminal (with `(.venv)` showing at the start of the prompt):
   ```
   python radOrderAutomate.py
   ```
4. The script validates your spreadsheet first and will stop with a clear error list if anything's wrong before it ever touches VistA.
5. You'll get a 5-second countdown to switch focus to your VistA terminal window before typing begins.
6. Watch the first order closely. If a `[HANDOFF]` message appears (with a beep), read it, fix the issue manually in VistA, then press **ESC** to resume.

### Testing checklist before a real run

- [ ] `ORDER_PREFIX` is set to a clearly fake value (e.g. `"ZZTEST "`) — either in your test account, or in production if your test account's procedure data is out of date (see note above)
- [ ] You're already sitting at the `Select QUICK ORDER NAME:` prompt in VistA before starting the script
- [ ] Spreadsheet headers are exactly `order_name` and `imaging_type`
- [ ] Clinical History template is on your clipboard
- [ ] You've watched at least one order complete successfully end-to-end
- [ ] You know what a `[HANDOFF]` pause looks like and how to resume with ESC

---

## Configuration Reference

All settings live in the `CONFIG` section at the top of `radOrderAutomate.py`.

| Setting | Purpose |
|---|---|
| `EXCEL_PATH` | Full path to your order spreadsheet. |
| `DRY_RUN` | `True` = print what would happen without sending real keystrokes. Good for a first check. |
| `TYPE_PAUSE` | Seconds between each simulated keystroke. `0.09` has proven reliable — going much lower risks corrupted keystrokes (see [Process Notes](#process-notes--lessons-learned)). |
| `PASTE_DELAY` | Initial pause after pasting Clinical History, before polling for paste completion. |
| `ORDER_PREFIX` / `ORDER_SUFFIX` | Your site's naming convention (e.g. `RAZN ... 2026`). **Set `ORDER_PREFIX` to an obviously-fake value (e.g. `"ZZTEST "`) while testing**, and only change it to your real production prefix once you trust the script's behavior. |
| `VALID_IMAGING_TYPES` | The set of imaging types the pre-flight check will accept — matches VistA's actual menu options. |

---

## Troubleshooting

**Script hangs for a long time right after starting, then hands off mentioning "Select QUICK ORDER NAME:"** — This almost always means VistA wasn't sitting at that exact prompt when the script started. Navigate to the Quick Order entry menu in VistA first, confirm `Select QUICK ORDER NAME:` is visible on screen, and re-run the script.

**Procedure name gets rejected — check which cause first.** There are two different reasons this can happen, and the fix depends on which one it is:
1. **Radiology hasn't added the procedure into VistA yet at all** (in any environment). This is a hard prerequisite — the script cannot work around it. Confirm with radiology whether the procedure has actually been added before assuming anything else is wrong.
2. **The procedure exists in production but your test account is out of date** (see note above). In this case the procedure is genuinely ready to use, just not yet visible in test.

Don't assume it's a script bug until both of these have been ruled out.

**`ModuleNotFoundError` for any package** — A required package isn't installed in the active virtual environment. In VS Code's terminal, with `(.venv)` showing, re-run:
```
pip install pywin32 pywinauto pyperclip pandas openpyxl
```
Virtual environments are isolated — packages installed elsewhere (or in a different environment) won't be available here even if you've installed them before.

**Spreadsheet rejected before the run starts** — The pre-flight validation will name the exact row and issue (wrong imaging type spelling, name too long, duplicate name, stray whitespace). Fix the row and re-run.

**`[HANDOFF]` pause with a beep** — The script didn't see the VistA prompt it expected. Read the message, fix the issue by hand in VistA, then press **ESC** to resume.

**`[TYPO DETECTED]` appears** — The script noticed a mismatch between what it typed and what's on screen (a rare keystroke timing glitch). It automatically clears and retypes the field. If it fails twice, it hands off to you.

**Order name / display text / procedure gets corrupted with a repeated letter** (e.g. "CONTRAST" → "CONTRRRR") — This was a known intermittent issue at faster typing speeds. Increasing `TYPE_PAUSE` and relying on the built-in verify/retry logic (already present in the script) resolves it. If you see this and the script doesn't catch it, that's a bug — please report it.

**A field gets skipped or double-filled (e.g. an order name typed into "Reason for Study")** — Some imaging types (e.g. Mammography, General Radiology) don't always show the same downstream prompts as others (some skip `Procedure Modifier:` entirely). If you hit a case like this, check whether the affected VistA order needs manual correction, and report the exact screen output so the step-detection logic can be updated for that case.

---

## Process Notes & Lessons Learned

### Background

This automation was built to support a HICOP requiring a large batch of new radiology quick orders — nearly all new orders on all new menus, aside from a handful of CT orders already built for an existing menu. Building these manually, one field at a time, would have taken months and tied up multiple staff in repetitive data entry, so automating the bulk of the process became a priority.

### Coordinating with Radiology

A shared spreadsheet was used to track progress between CACs and radiology staff:

- **Orange highlight** — quick order created.
- **Green highlight** — order not done at this site.
- **Grey highlight** — "Tech Order Only," no order creation required.
- **No highlight** — not yet completed by radiology.

Once radiology marked a study "Activated in Rad Package" and "Radiology testing complete," CACs knew it was ready for quick order creation. These two markers matter because they represent the actual hard prerequisite for this entire process: **the script can only create a quick order for a procedure that already exists in VistA.** Attempting an order before radiology has added and activated its procedure will fail regardless of how the automation is configured — this isn't a script limitation to work around, it's a sequencing dependency that has to be respected. Completed orders had their exact quick order name recorded for a clear audit trail.

### First Approach: Macros

Before building a script, we looked into VistA macros. Our understanding at the time was that macros only handled menu navigation, not the multi-step, conditional process of order creation — this later turned out to be incorrect, based on an example seen elsewhere. Given that we didn't know this at the time, and had a stronger existing background in Python, we built a custom script instead — which also gave more control over validating VistA's output at each step, which proved essential.

### Building the Automation

Built with `pywinauto` to simulate keystrokes and `win32com` to read the terminal screen back for confirmation, against a Reflection Desktop Pro / EXTRA! session. The connection code built for EXTRA! also worked unmodified against Reflection Desktop Pro, which matters for sharing this across sites using either emulator.

**Unpredictable VistA output.** Early versions assumed a fixed prompt sequence, which broke down quickly given VistA's variable output (prep text, category warnings, disambiguation prompts). The fix: wait for explicit screen confirmation before every typed action, rather than assuming fixed timing.

**Naming convention.** A `2026` suffix was added to every new order both to differentiate them from older orders and guarantee uniqueness. The prefix and suffix together are what make each order name unique — this matters because VistA requires each quick order dialog name to be distinct, and attempting to "create" a name that already exists means editing that existing order instead of creating a new one. We kept our normal production prefix (`RAZN`) rather than inventing a new one, and paired it with a suffix (`2026`) that hadn't been used before, so that no existing order name could collide with a newly created one. This is a naming-convention safeguard, separate from (and in addition to) the script's own pre-flight duplicate check described below — even if a naming collision were to happen despite the convention, the script's validation step would catch and stop it before the run ever starts, rather than silently overwriting an existing order. A swappable test prefix (`ZZTEST` or similar) made test-run orders easy to distinguish from production-track builds.

**Order name length limit.** An initial assumption of a 30-character limit was wrong. Empirical testing (typing progressively longer strings until VistA rejected one) found the real limit to be 64 characters, which is now enforced as a pre-flight check.

**CPT codes vs. procedure names.** Radiology's master spreadsheet paired each scan with a CPT code. Typing the CPT code directly works only when it maps to exactly one scan — many codes map to multiple (e.g. with/without contrast, or left/right pairs), making CPT entry ambiguous. Typing the full procedure name directly proved reliable in every case, so the spreadsheet was simplified to just scan name and imaging type.

**Redundant spreadsheet column.** Since display name and procedure name were always identical by convention, the spreadsheet was simplified from three columns to two, removing a place where the two values could silently drift out of sync.

**Procedure name truncation.** VistA matches typed procedure names on only their first 30 characters. Names longer than that which share a prefix with another (e.g. "...CONTRAST LEFT" vs "...CONTRAST RIGHT") trigger a numbered disambiguation prompt instead of being accepted directly. The script detects this, parses the numbered list, and matches the full untruncated name against it to select automatically — only handing off if no exact match is found.

**Special characters.** Characters like `+` are treated as modifier-key syntax by the keystroke-simulation library and were being silently dropped or misinterpreted. Fixed by explicitly escaping this small set of reserved characters before typing.

**Clinical History formatting.** Typing the clinical history template character-by-character never reproduced correct formatting. The fix was pasting a pre-set clipboard template instead. Paste timing itself needed two iterations: a fixed delay worked for small templates but caused spillover into the next field for larger ones; the eventual fix polls the screen until its content stops changing, confirming the paste has genuinely finished before proceeding.

**Typing speed and stuck keys.** At a fast typing pace, a single character would occasionally repeat several times in a row, corrupting a field (e.g. "CONTRAST" → "CONTRRRR"). Slowing the pace helped but didn't eliminate it, and the bug was too rare to reliably reproduce for testing — so a temporary debug switch was built to force the failure path on demand and confirm the fix worked. This process also revealed that even short, fixed strings could be affected, disproving the assumption that only long values were at risk. Verification (type → read back → retry on mismatch → hand off if still wrong) was extended to every typed field as a result.

**Inconsistent downstream prompts.** Not every imaging type shows the same sequence of prompts after the procedure name — Mammography, for instance, can skip `Procedure Modifier:` entirely and go straight to `Reason for Study:`. Detection logic had to be updated to treat multiple possible "acceptance" signals as valid, and to avoid false-positive matches against stale text still visible on screen from a previous prompt.

**Environment setup friction.** Fresh virtual environments don't inherit previously installed packages, even on a machine where they'd been used before — a predictable first hurdle worth documenting explicitly rather than assuming away.

**Pre-flight validation.** As batch sizes grew, so did the cost of discovering a bad row mid-run. A validation pass now checks the entire spreadsheet upfront — imaging type spelling, name length, duplicate names, stray whitespace — before the script ever touches the VistA session.

### Summary of Key Lessons

- Don't assume a fixed sequence of prompts — confirm actual on-screen state at every step.
- Naming conventions should serve automation as well as humans.
- CPT codes aren't always unambiguous identifiers; matching on full procedure names is more reliable, though VistA's own 30-character truncation adds a separate ambiguity to handle.
- Clipboard paste beats typed text for formatted content, but paste timing needs active confirmation, not a guessed delay.
- Keystroke-simulation tools can silently misinterpret special characters — escape anything typed programmatically.
- Don't assume risk is proportional to string length — verify everything typed, not just what looks risky.
- When a bug is too rare to observe reliably, build a way to force it on demand.
- Validate data before a run starts, not during it.
- Isolated environments require explicit, repeatable setup steps.
