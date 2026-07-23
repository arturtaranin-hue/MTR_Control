# Building `MTR_Control.xlsm`

`MTR_Control.xlsm` is a generated release artifact. It is intentionally not stored in this source repository or pull request because macro-enabled Excel workbooks are binary files. Build it locally from the tracked VBA source files when you need a runnable workbook.

## Prerequisites

1. Windows with Microsoft Excel installed.
2. PowerShell 5.1 or newer.
3. Excel configured to allow VBA project automation:
   - Open Excel.
   - Go to **File > Options > Trust Center > Trust Center Settings > Macro Settings**.
   - Enable **Trust access to the VBA project object model**.
4. Macro settings appropriate for your organization so you can save and run a local `.xlsm` workbook.

## Build steps

From the repository root, run:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\Build-MTRControl.ps1
```

The script will:

1. Create `releases\MTR_Control.xlsm`.
2. Import every tracked VBA source file from `src\` into the workbook, including `.bas`, `.cls`, and `.frm` files.
3. Run `SetupDashboard` from `src\modMvp.bas` to create the `Dashboard` worksheet and add the buttons:
   - **Import Files**
   - **Generate Report**
   - **Open Log**
4. Save and close the generated macro-enabled workbook.

## Custom output path

To build to another location, pass `-OutputPath`:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\Build-MTRControl.ps1 -OutputPath C:\Temp\MTR_Control.xlsm
```

## Running the MVP

1. Open the generated workbook in Excel.
2. Enable macros if prompted.
3. On the `Dashboard` worksheet, click **Import Files** or **Generate Report**.
4. Select the three source files when prompted:
   - DV Angara
   - Movement Summary
   - GPN Angara
5. Review the generated `Report`, `Dashboard`, and `Log` worksheets.

## Troubleshooting

- If PowerShell reports that VBA project access is denied, enable **Trust access to the VBA project object model** in Excel and rerun the script.
- If Excel is already open with a workbook of the same output name, close that workbook before building.
- If your organization blocks macros, build the workbook in an approved trusted location or follow your internal macro-signing process.
