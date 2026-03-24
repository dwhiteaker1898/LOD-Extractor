import csv
import os
import sys
import traceback
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


SERIES_OPTIONS = ("ALOD", "CMLOD")
VIEW_OPTIONS = (
    ("hourly", "Hourly"),
    ("daily", "Daily"),
    ("monthly", "Monthly"),
    ("annually", "Annually"),
)
EXPECTED_FIELD_COUNT = 178
METADATA_FIELD_COUNT = 10
HOURS_PER_WEEK = 168
SUCCESS_PREFIX = "\u2713"


def application_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def default_output_dir(source_folder: str) -> str:
    return source_folder


@dataclass
class LodRecord:
    file_name: str
    week_start: datetime.date
    area: str
    series_type: str
    values: list[float]
    start_hour_offset: int = 0


def parse_lod_file(path: str, series_type: str, trim_edge_weeks: bool) -> list[LodRecord]:
    records: list[LodRecord] = []

    with open(path, "r", newline="") as handle:
        for line_number, raw_line in enumerate(handle, start=1):
            line = raw_line.strip()
            if not line:
                continue

            parts = [part.strip() for part in line.split(",")]
            if len(parts) != EXPECTED_FIELD_COUNT:
                raise ValueError(
                    f"{os.path.basename(path)} line {line_number}: expected "
                    f"{EXPECTED_FIELD_COUNT} fields, found {len(parts)}."
                )

            current_series = parts[5]
            if current_series != series_type:
                continue

            week_start = datetime.strptime(parts[1], "%m/%d/%Y").date()
            values = [float(value) for value in parts[METADATA_FIELD_COUNT:]]

            if len(values) != HOURS_PER_WEEK:
                raise ValueError(
                    f"{os.path.basename(path)} line {line_number}: expected "
                    f"{HOURS_PER_WEEK} hourly values, found {len(values)}."
                )

            records.append(
                LodRecord(
                    file_name=os.path.basename(path),
                    week_start=week_start,
                    area=parts[3],
                    series_type=current_series,
                    start_hour_offset=0,
                    values=values,
                )
            )

    if not records:
        raise ValueError(f"{os.path.basename(path)} does not contain any {series_type} records.")

    if trim_edge_weeks:
        hour_counts: Counter[int] = Counter()
        for record in records:
            start_dt = datetime.combine(record.week_start, datetime.min.time())
            for offset in range(len(record.values)):
                timestamp = start_dt + timedelta(hours=offset)
                hour_counts[timestamp.year] += 1

        dominant_year = hour_counts.most_common(1)[0][0]
        trimmed_records: list[LodRecord] = []
        for record in records:
            start_dt = datetime.combine(record.week_start, datetime.min.time())
            included_offsets = [
                offset
                for offset in range(len(record.values))
                if (start_dt + timedelta(hours=offset)).year == dominant_year
            ]
            if not included_offsets:
                continue

            first_offset = included_offsets[0]
            last_offset = included_offsets[-1]
            trimmed_records.append(
                LodRecord(
                    file_name=record.file_name,
                    week_start=record.week_start,
                    area=record.area,
                    series_type=record.series_type,
                    start_hour_offset=record.start_hour_offset + first_offset,
                    values=record.values[first_offset : last_offset + 1],
                )
            )

        records = trimmed_records

    return records


def hourly_rows(records: list[LodRecord]) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []

    for record in records:
        start_dt = datetime.combine(record.week_start, datetime.min.time()) + timedelta(
            hours=record.start_hour_offset
        )
        for offset, value in enumerate(record.values):
            timestamp = start_dt + timedelta(hours=offset)
            rows.append(
                {
                    "source_file": record.file_name,
                    "area": record.area,
                    "series_type": record.series_type,
                    "timestamp": timestamp.strftime("%Y-%m-%d %H:%M:%S"),
                    "year": timestamp.year,
                    "month": timestamp.month,
                    "day": timestamp.day,
                    "hour": timestamp.hour + 1,
                    "energy_mwh": round(value, 6),
                }
            )

    return rows


def aggregate_rows(records: list[LodRecord], view_type: str) -> list[dict[str, object]]:
    grouped: dict[tuple, dict[str, float]] = {}

    for record in records:
        start_dt = datetime.combine(record.week_start, datetime.min.time()) + timedelta(
            hours=record.start_hour_offset
        )
        for offset, value in enumerate(record.values):
            timestamp = start_dt + timedelta(hours=offset)

            if view_type == "daily":
                key = (
                    record.file_name,
                    record.area,
                    record.series_type,
                    timestamp.year,
                    timestamp.month,
                    timestamp.day,
                )
            elif view_type == "monthly":
                key = (
                    record.file_name,
                    record.area,
                    record.series_type,
                    timestamp.year,
                    timestamp.month,
                )
            elif view_type == "annually":
                key = (
                    record.file_name,
                    record.area,
                    record.series_type,
                    timestamp.year,
                )
            else:
                raise ValueError(f"Unsupported view type: {view_type}")

            stats = grouped.setdefault(key, {"total_mwh": 0.0, "peak_mw": float("-inf")})
            stats["total_mwh"] += value
            stats["peak_mw"] = max(stats["peak_mw"], value)

    rows: list[dict[str, object]] = []
    for key in sorted(grouped):
        stats = grouped[key]
        total = round(stats["total_mwh"], 6)
        peak = round(stats["peak_mw"], 6)

        if view_type == "daily":
            source_file, area, series_type, year, month, day = key
            rows.append(
                {
                    "source_file": source_file,
                    "area": area,
                    "series_type": series_type,
                    "year": year,
                    "month": month,
                    "day": day,
                    "energy_mwh": total,
                    "peak_mw": peak,
                }
            )
        elif view_type == "monthly":
            source_file, area, series_type, year, month = key
            rows.append(
                {
                    "source_file": source_file,
                    "area": area,
                    "series_type": series_type,
                    "year": year,
                    "month": month,
                    "energy_mwh": total,
                    "peak_mw": peak,
                }
            )
        elif view_type == "annually":
            source_file, area, series_type, year = key
            rows.append(
                {
                    "source_file": source_file,
                    "area": area,
                    "series_type": series_type,
                    "year": year,
                    "energy_mwh": total,
                    "peak_mw": peak,
                }
            )

    return rows


def summarize_records(records: list[LodRecord], view_type: str) -> list[dict[str, object]]:
    if view_type == "hourly":
        return hourly_rows(records)
    return aggregate_rows(records, view_type)


def write_csv(path: str, rows: list[dict[str, object]]) -> None:
    if not rows:
        raise ValueError("No rows were generated for export.")

    fieldnames = list(rows[0].keys())
    with open(path, "w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


class LodSummaryApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("PROMOD LOD Summary Export")
        self.root.geometry("900x640")
        self.root.minsize(900, 640)

        self.selected_folder = tk.StringVar(value="")
        self.selected_files: list[str] = []
        self.output_dir = tk.StringVar(value="")
        self.series_type = tk.StringVar(value="ALOD")
        self.view_vars = {key: tk.BooleanVar(value=(key == "hourly")) for key, _ in VIEW_OPTIONS}
        self.trim_edge_weeks = tk.BooleanVar(value=True)
        self.combine_output = tk.BooleanVar(value=False)
        self.status_text = tk.StringVar(value="Select a folder that contains .LOD files, then export.")

        self._build_ui()

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=16)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)

        file_frame = ttk.LabelFrame(container, text="Source Folder", padding=12)
        file_frame.grid(row=0, column=0, sticky="nsew")
        file_frame.columnconfigure(0, weight=1)

        file_button_row = ttk.Frame(file_frame)
        file_button_row.grid(row=0, column=0, sticky="ew")
        file_button_row.columnconfigure(0, weight=1)

        ttk.Entry(file_button_row, textvariable=self.selected_folder).grid(row=0, column=0, sticky="ew")
        ttk.Button(file_button_row, text="Select Source Folder", command=self.select_folder).grid(
            row=0, column=1, padx=(8, 0)
        )

        self.file_list = tk.Listbox(file_frame, height=12)
        self.file_list.grid(row=1, column=0, sticky="nsew", pady=(10, 0))

        options_frame = ttk.LabelFrame(container, text="Export Options", padding=12)
        options_frame.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        options_frame.columnconfigure(0, weight=1)
        options_frame.columnconfigure(1, weight=1)

        series_frame = ttk.Frame(options_frame)
        series_frame.grid(row=0, column=0, sticky="nw")
        ttk.Label(series_frame, text="Series").grid(row=0, column=0, sticky="w")
        for index, option in enumerate(SERIES_OPTIONS, start=1):
            ttk.Radiobutton(series_frame, text=option, value=option, variable=self.series_type).grid(
                row=index, column=0, sticky="w"
            )

        view_frame = ttk.Frame(options_frame)
        view_frame.grid(row=0, column=1, sticky="nw")
        ttk.Label(view_frame, text="Reports").grid(row=0, column=0, sticky="w")
        for index, (key, label) in enumerate(VIEW_OPTIONS, start=1):
            ttk.Checkbutton(view_frame, text=label, variable=self.view_vars[key]).grid(
                row=index, column=0, sticky="w"
            )

        ttk.Checkbutton(options_frame, text="Trim edge weeks", variable=self.trim_edge_weeks).grid(
            row=1, column=0, sticky="w", pady=(12, 0)
        )
        ttk.Checkbutton(
            options_frame,
            text="Combine Into One Report",
            variable=self.combine_output,
        ).grid(row=1, column=1, sticky="w", pady=(12, 0))

        output_frame = ttk.LabelFrame(container, text="Output Folder", padding=12)
        output_frame.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        output_frame.columnconfigure(0, weight=1)

        ttk.Entry(output_frame, textvariable=self.output_dir).grid(row=0, column=0, sticky="ew")
        ttk.Button(output_frame, text="Browse", command=self.select_output_dir).grid(row=0, column=1, padx=(8, 0))

        action_frame = ttk.Frame(container)
        action_frame.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        action_frame.columnconfigure(0, weight=1)

        ttk.Button(action_frame, text="Export CSV Files", command=self.export_files).grid(row=0, column=0, sticky="w")
        ttk.Label(action_frame, textvariable=self.status_text).grid(row=0, column=1, sticky="e")

        self._refresh_file_list()

    def select_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select folder containing LOD files")
        if not folder:
            return

        self.selected_folder.set(folder)
        self.output_dir.set(default_output_dir(folder))
        self.refresh_files_from_folder()

    def selected_views(self) -> list[str]:
        return [key for key, _ in VIEW_OPTIONS if self.view_vars[key].get()]

    def refresh_files_from_folder(self) -> None:
        folder = self.selected_folder.get().strip()
        if not folder:
            self.selected_files = []
            self._refresh_file_list()
            self.status_text.set("No source folder selected.")
            return

        if not os.path.isdir(folder):
            messagebox.showerror("Invalid folder", "Select a valid folder containing .LOD files.")
            return

        self.selected_files = sorted(
            os.path.join(folder, name)
            for name in os.listdir(folder)
            if name.lower().endswith('.lod')
        )
        self._refresh_file_list()
        if self.selected_files:
            self.status_text.set(f"Found {len(self.selected_files)} .LOD file(s) in folder.")
        else:
            self.status_text.set("No .LOD files found in source folder.")

    def _refresh_file_list(self) -> None:
        self.file_list.delete(0, tk.END)
        if self.selected_files:
            for path in self.selected_files:
                self.file_list.insert(tk.END, path)
        else:
            self.file_list.insert(tk.END, "No .LOD files found in source folder")

    def select_output_dir(self) -> None:
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self.output_dir.set(folder)

    def export_files(self) -> None:
        if not self.selected_files:
            messagebox.showerror("No files selected", "Select a folder that contains at least one .LOD file.")
            return

        selected_views = self.selected_views()
        if not selected_views:
            messagebox.showerror("No reports selected", "Select at least one report type to export.")
            return

        output_dir = self.output_dir.get().strip()
        if not output_dir:
            messagebox.showerror("No output folder", "Select an output folder.")
            return

        if os.path.abspath(output_dir) == os.path.abspath(application_dir()):
            messagebox.showerror(
                "Invalid output folder",
                "Choose an output folder other than the application folder.",
            )
            return

        os.makedirs(output_dir, exist_ok=True)
        exported_files: list[str] = []
        failures: list[str] = []

        self.status_text.set("Exporting...")
        self.root.update_idletasks()

        parsed_records_by_file: dict[str, list[LodRecord]] = {}
        for path in self.selected_files:
            try:
                parsed_records_by_file[path] = parse_lod_file(
                    path=path,
                    series_type=self.series_type.get(),
                    trim_edge_weeks=self.trim_edge_weeks.get(),
                )
            except Exception as exc:
                failures.append(f"{os.path.basename(path)}: {exc}")

        if self.combine_output.get():
            all_records: list[LodRecord] = []
            for path in self.selected_files:
                all_records.extend(parsed_records_by_file.get(path, []))

            if all_records:
                for view_type in selected_views:
                    rows = summarize_records(all_records, view_type)
                    view_label = dict(VIEW_OPTIONS)[view_type]
                    output_name = f"Combined_Load_{view_label}.csv"
                    output_path = os.path.join(output_dir, output_name)

                    write_csv(output_path, rows)
                    exported_files.append(output_path)
        else:
            for path in self.selected_files:
                records = parsed_records_by_file.get(path)
                if not records:
                    continue

                base_name = os.path.splitext(os.path.basename(path))[0]
                for view_type in selected_views:
                    rows = summarize_records(records, view_type)
                    view_label = dict(VIEW_OPTIONS)[view_type]
                    output_name = f"{base_name}_Load_{view_label}.csv"
                    output_path = os.path.join(output_dir, output_name)

                    write_csv(output_path, rows)
                    exported_files.append(output_path)

        if not exported_files:
            self.status_text.set("No files exported.")
            messagebox.showerror(
                "Export failed",
                "No CSV files were created. Review the errors and try again.",
            )
            return

        if failures:
            self.status_text.set("Completed with errors.")
            messagebox.showwarning(
                "Export completed with errors",
                "Some files were not exported:\n\n" + "\n".join(failures),
            )
        else:
            self.status_text.set(f"Exported {len(exported_files)} file(s).")
            messagebox.showinfo(
                "Export complete",
                f"{SUCCESS_PREFIX} Export complete. Created CSV files:\n\n" + "\n".join(exported_files),
            )
            self.root.destroy()


def report_fatal_error(exc: Exception) -> None:
    error_text = "".join(traceback.format_exception(exc))
    try:
        messagebox.showerror("Unexpected error", error_text)
    except Exception:
        pass
    print(error_text, file=sys.stderr)


def main() -> None:
    root = tk.Tk()
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")

    LodSummaryApp(root)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        report_fatal_error(exc)
        raise
