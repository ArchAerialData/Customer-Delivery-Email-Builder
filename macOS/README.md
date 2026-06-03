# macOS Support

These files are for running the app on a Mac as a Customer Information workbook editor.

## Setup

Install Python 3 from python.org, then install the app dependencies:

```bash
python3 -m pip install -r macOS/requirements-macos.txt
```

## Run

Open `Run Email Builder.command`, or run this from the project folder:

```bash
python3 email_template_gui.py
```

If macOS says the command file is not executable, run:

```bash
chmod +x "macOS/Run Email Builder.command"
```

On macOS, email draft creation is disabled. The workbook editing screens still use `Reference Data/XLSX Workbooks/Customer Information.xlsx`.
