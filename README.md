# PROMOD LOD Summary Tool

This tool reads PROMOD `.LOD` files from a selected folder and exports normalized CSV summaries.

## Features

- Desktop popup built with Tkinter
- Select a folder and scan all `.LOD` files in it
- Choose `ALOD` or `CMLOD`
- Choose `hourly`, `daily`, `monthly`, or `annually`
- Export one CSV per `.LOD` file found in the folder
- Output names follow the pattern `OriginalFileName_Load_Monthly.csv`
- Optional `Trim edge weeks` behavior to remove boundary weeks outside the main study year

## Output behavior

- `hourly`: one row per hour with `year`, `month`, `day`, and `hour`
- `daily`: sums all hourly values within each day
- `monthly`: sums all hourly values within each month
- `annually`: sums all hourly values within each year
- Output always includes all areas found in each file

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
