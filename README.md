# Makro Solver

A small Windows tool for Excel files you can open but can no longer edit because
you have forgotten a worksheet, workbook, or VBA project password.

Put your files in `input`, run the tool, and pick up the unprotected copies from
`output`. Your originals stay where they are. Everything runs locally, without
opening Excel or executing any macros.

This tool removes protection settings. It does **not** recover your old password
or decrypt a file that needs a password before it can be opened.

## Getting started

You need Windows and Python **3.14**, including the Python Launcher (`py`).
The required `olefile` package is included in this folder, so there is nothing
else to install. Excel and an internet connection are not needed to run the tool.

1. Download or copy the whole project folder.
2. Put your `.xlsb`, `.xlsm`, or `.xlsx` files directly inside `input`.
3. Double-click **START.bat**.
4. Open `output` to find your copies and a JSON report of the run.

If `input` and `output` are missing after downloading the project, run the tool
once. It will create both folders and ask you to add your files.

Already have an output with the same name? The next copy gets a suffix such as
`(2)` or `(3)`. A file that cannot be processed is listed as an error, and the tool
continues with the rest. Inputs are left in place, so running it again processes
them again. It does not watch the folder in the background or scan subfolders.

## What it handles

| Protection | Supported? |
| --- | --- |
| Worksheet editing protection | Yes |
| Workbook structure protection | Yes |
| VBA project viewing password | Yes, when a VBA project is present |
| Password required to open an encrypted file | No |
| Password prompts implemented by a macro | No |
| Older `.xls` files | No |
| Digitally signed packages | No; editing could invalidate the signature |

Macro code is preserved. A macro that protects a sheet when it runs can still
do that afterwards. Check the output in Excel before relying on it, then set a
new password there if you want one.

## Running it from a terminal

From the project folder:

```powershell
py -3.14 solver.py
```

For a scheduled task, pass the full path to `solver.py`. The input and output
folders are always resolved relative to the script, regardless of the terminal's
working directory. Use the Python command rather than `START.bat` for unattended
runs, since the batch launcher waits for a keypress.

Exit code `0` means the run succeeded or the input folder was empty. Exit code `1`
means at least one file failed or a program-level error occurred.

The launcher selects Python 3.14. The code uses `hashlib.file_digest`, which needs
Python 3.11 or newer; **Python 3.9 is not supported yet**. Testing has been done
on Python 3.14. Do not use `python -O`, because it disables validation assertions.

## What happens inside

Excel files in these formats are packages containing workbook data and, sometimes,
a VBA project. The tool edits the relevant protection settings in a temporary
copy, then checks that copy before giving it its final filename.

- For `.xlsb`, it removes the worksheet and workbook protection records from
  the binary streams, keeping all other records as they were.
- For `.xlsx` and `.xlsm`, it removes the corresponding protection elements from
  the XML without rewriting the rest of the document.
- For VBA, it updates `CMG`, `DPB`, and `GC` in the `PROJECT` stream to represent
  an unlocked, visible project with no password. Other OLE streams are compared
  byte for byte, including the streams containing code and forms.

It also checks ZIP integrity, compares unchanged package parts, and verifies that
the input file's SHA-256 hash has not changed during processing. These are file
checks; they do not prove that every workbook feature or macro will work in Excel.

## Finding your way around the code

| File | Purpose |
| --- | --- |
| `solver.py` | Batch processing, XML changes, output validation, and reports |
| `protection.py` | Binary record parsing and VBA protection settings |
| `test_solver.py` | Regression tests for processing and output handling |
| `START.bat` | Double-click launcher for Windows |
| `USAGE.txt` | Short instructions to keep alongside the program |

New reports use English field names, such as `input`, `status`, `output`, and
`modified_parts`. Protection counts refer to records, not necessarily individual
sheets. `vba_processed` means the VBA protection fields were processed; it does
not mean the macros were executed or tested. Existing reports are left untouched.

## Tests

```powershell
py -3.14 -m unittest -v test_solver
```

The tests cover XML protection removal, continuing after a damaged file, preserving
the input, and choosing a new name when an output already exists.

An additional integration test uses a local `.xlsb` sample from a sibling folder
named `bbsk rapor`. It is skipped when that folder is absent. Private workbooks
are not included as public test fixtures. The current suite does not open Excel
or test macro execution.

## Before publishing or sharing

The `.gitignore` excludes `input`, `output`, Excel files, and Python caches.
It will not remove files that Git already tracks. Check your staged files before
committing. If you share a ZIP instead, leave your personal input and output files
out of it; `.gitignore` does not filter a ZIP you create manually.

The tool currently targets Windows. Its output naming relies on Windows refusing
to rename a file over an existing destination; do not assume the same behavior
on Linux or macOS.

The bundled `olefile` dependency includes its own license information in
`olefile-0.47.dist-info`. Keep those notices when redistributing it. A license for
this project's own code has not been chosen yet.
