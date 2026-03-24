# PROMOD LOD Summary Tool

This tool reads PROMOD `.LOD` files from a selected folder and exports normalized CSV summaries.

## Features

- Desktop popup built with Tkinter
- Select a folder and scan all `.LOD` files in it
- Choose `ALOD` or `CMLOD`
- Choose `hourly`, `daily`, `monthly`, or `annually`
- Export one CSV per `.LOD` file found in the folder
- Optionally combine all successfully parsed `.LOD` files into one CSV per selected report
- Output names follow the pattern `OriginalFileName_Load_Monthly.csv`
- Optional `Trim edge weeks` behavior to keep only the dominant study year's hourly data, including leap-year hours when present

## Output behavior

- `hourly`: one row per hour with `year`, `month`, `day`, and `hour`
- `daily`: sums all hourly values within each day and records the interval peak load
- `monthly`: sums all hourly values within each month and records the interval peak load
- `annually`: sums all hourly values within each year and records the interval peak load
- Output always includes all areas found in each file
- Combined exports retain the `source_file` column so each row can still be traced to its original `.LOD` file
- Energy totals are exported as `energy_mwh` and peak demand is exported as `peak_mw`

## Run with Python

```bat
py lod_summary_app.py
```

## Build an executable

If PyInstaller is installed:

```bat
build_exe.bat
```

The executable will be created at `dist\lod_summary_app.exe`.
