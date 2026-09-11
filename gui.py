import json
import re
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from comparator import compare_versions
from excel_handler import open_output_folder


SETTINGS_FILE = (
    Path(__file__).resolve().parent
    / "settings.json"
)


class ExcelVersionDiffApp:
    def __init__(self, root):
        self.root = root

        self.root.title(
            "Excel Version Difference Checker"
        )

        self.root.geometry(
            "800x650"
        )

        self.root.minsize(
            720,
            560,
        )

        self.old_folder_var = (
            tk.StringVar()
        )

        self.new_folder_var = (
            tk.StringVar()
        )

        self.output_var = (
            tk.StringVar(
                value="-"
            )
        )

        self.pairing_method_var = (
            tk.StringVar(
                value="filename"
            )
        )

        self.parallel_var = (
            tk.BooleanVar(
                value=False
            )
        )

        self.cancel_event = (
            threading.Event()
        )

        self.worker_thread = None

        self.compare_start_time = None

        self.progress_states = {}

        self.progress_total = 0

        self.load_settings()
        self.build_ui()

    def load_settings(self):
        if not SETTINGS_FILE.exists():
            return

        try:
            with open(
                SETTINGS_FILE,
                "r",
                encoding="utf-8",
            ) as f:
                settings = json.load(f)

            self.old_folder_var.set(
                settings.get(
                    "old_folder",
                    "",
                )
            )

            self.new_folder_var.set(
                settings.get(
                    "new_folder",
                    "",
                )
            )

            self.pairing_method_var.set(
                settings.get(
                    "pairing_method",
                    "filename",
                )
            )

            self.parallel_var.set(
                settings.get(
                    "parallel",
                    False,
                )
            )

        except (
            OSError,
            json.JSONDecodeError,
        ):
            pass

    def save_settings(self):
        settings = {
            "old_folder": (
                self.old_folder_var
                .get()
                .strip()
            ),
            "new_folder": (
                self.new_folder_var
                .get()
                .strip()
            ),
            "pairing_method": (
                self.pairing_method_var.get()
            ),
            "parallel": (
                self.parallel_var.get()
            ),
        }

        try:
            with open(
                SETTINGS_FILE,
                "w",
                encoding="utf-8",
            ) as f:
                json.dump(
                    settings,
                    f,
                    ensure_ascii=False,
                    indent=4,
                )

        except OSError:
            pass

    def build_ui(self):
        frame = ttk.Frame(
            self.root,
            padding=20,
        )

        frame.pack(
            fill="both",
            expand=True,
        )

        ttk.Label(
            frame,
            text=(
                "Excel Version "
                "Difference Checker"
            ),
            font=(
                "Segoe UI",
                16,
                "bold",
            ),
        ).pack(
            anchor="w",
            pady=(0, 20),
        )

        self.add_folder_row(
            frame,
            "Old Version Folder",
            self.old_folder_var,
        )

        self.add_folder_row(
            frame,
            "New Version Folder",
            self.new_folder_var,
        )

        pairing_frame = ttk.LabelFrame(
            frame,
            text="Pairing Method",
            padding=10,
        )

        pairing_frame.pack(
            fill="x",
            pady=(15, 10),
        )

        ttk.Radiobutton(
            pairing_frame,
            text="Filename",
            variable=(
                self.pairing_method_var
            ),
            value="filename",
            command=self.save_settings,
        ).pack(
            side="left",
            padx=(0, 20),
        )

        ttk.Radiobutton(
            pairing_frame,
            text="Folder Order",
            variable=(
                self.pairing_method_var
            ),
            value="order",
            command=self.save_settings,
        ).pack(
            side="left",
        )

        parallel_frame = ttk.LabelFrame(
            frame,
            text="Parallel Processing",
            padding=10,
        )

        parallel_frame.pack(
            fill="x",
            pady=(0, 10),
        )

        ttk.Radiobutton(
            parallel_frame,
            text="Off",
            variable=self.parallel_var,
            value=False,
            command=self.save_settings,
        ).pack(
            side="left",
            padx=(0, 20),
        )

        ttk.Radiobutton(
            parallel_frame,
            text="On",
            variable=self.parallel_var,
            value=True,
            command=self.save_settings,
        ).pack(
            side="left",
        )

        output_frame = ttk.Frame(
            frame
        )

        output_frame.pack(
            fill="x",
            pady=(5, 15),
        )

        ttk.Label(
            output_frame,
            text="Output",
        ).pack(
            anchor="w"
        )

        ttk.Label(
            output_frame,
            textvariable=self.output_var,
            relief="sunken",
            anchor="w",
        ).pack(
            fill="x",
            pady=(4, 0),
        )

        button_frame = ttk.Frame(
            frame
        )

        button_frame.pack(
            anchor="e",
            pady=(0, 15),
        )

        self.compare_button = ttk.Button(
            button_frame,
            text="Compare",
            command=self.start_compare,
        )

        self.compare_button.pack(
            side="left",
            padx=(0, 8),
        )

        self.stop_button = ttk.Button(
            button_frame,
            text="Stop",
            command=self.stop_compare,
            state="disabled",
        )

        self.stop_button.pack(
            side="left",
        )

        ttk.Label(
            frame,
            text="Log",
        ).pack(
            anchor="w"
        )

        log_frame = ttk.Frame(
            frame
        )

        log_frame.pack(
            fill="both",
            expand=True,
            pady=(4, 0),
        )

        self.log_text = tk.Text(
            log_frame,
            wrap="word",
            state="disabled",
        )

        scrollbar = ttk.Scrollbar(
            log_frame,
            orient="vertical",
            command=self.log_text.yview,
        )

        self.log_text.configure(
            yscrollcommand=scrollbar.set
        )

        self.log_text.pack(
            side="left",
            fill="both",
            expand=True,
        )

        scrollbar.pack(
            side="right",
            fill="y",
        )

    def add_folder_row(
        self,
        parent,
        label_text,
        variable,
    ):
        row = ttk.Frame(
            parent
        )

        row.pack(
            fill="x",
            pady=5,
        )

        ttk.Label(
            row,
            text=label_text,
            width=20,
        ).pack(
            side="left"
        )

        ttk.Entry(
            row,
            textvariable=variable,
        ).pack(
            side="left",
            fill="x",
            expand=True,
            padx=8,
        )

        ttk.Button(
            row,
            text="Browse",
            command=lambda: (
                self.select_folder(
                    variable
                )
            ),
        ).pack(
            side="right"
        )

    def select_folder(
        self,
        variable,
    ):
        folder = (
            filedialog.askdirectory()
        )

        if folder:
            variable.set(folder)
            self.save_settings()

    def clear_log(self):
        self.log_text.configure(
            state="normal"
        )

        self.log_text.delete(
            "1.0",
            "end",
        )

        self.log_text.configure(
            state="disabled"
        )

    def append_log(self, message):
        self.log_text.configure(
            state="normal"
        )

        self.log_text.insert(
            "end",
            message + "\n",
        )

        self.log_text.see(
            "end"
        )

        self.log_text.configure(
            state="disabled"
        )

    def log(self, message):
        self.root.after(
            0,
            lambda message=message: (
                self.append_log(message)
            ),
        )

    def start_compare(self):
        if (
            self.worker_thread is not None
            and self.worker_thread.is_alive()
        ):
            return

        old_folder = (
            self.old_folder_var
            .get()
            .strip()
        )

        new_folder = (
            self.new_folder_var
            .get()
            .strip()
        )

        if not old_folder:
            messagebox.showwarning(
                "Input Required",
                "Please select the Old Version folder.",
            )
            return

        if not new_folder:
            messagebox.showwarning(
                "Input Required",
                "Please select the New Version folder.",
            )
            return

        self.save_settings()

        pairing_method = (
            self.pairing_method_var.get()
        )

        parallel = (
            self.parallel_var.get()
        )

        self.cancel_event.clear()

        self.compare_start_time = (
            time.perf_counter()
        )

        self.progress_states = {}
        self.progress_total = 0

        self.compare_button.configure(
            state="disabled"
        )

        self.stop_button.configure(
            state="normal"
        )

        self.output_var.set("-")

        self.clear_log()

        pairing_name = (
            "Filename"
            if pairing_method == "filename"
            else "Folder Order"
        )

        parallel_name = (
            "On"
            if parallel
            else "Off"
        )

        self.append_log(
            "Starting comparison..."
        )

        self.append_log(
            f"Pairing Method: {pairing_name}"
        )

        self.append_log(
            f"Parallel Processing: {parallel_name}"
        )

        self.append_log(
            ""
        )

        self.worker_thread = (
            threading.Thread(
                target=self.run_compare,
                args=(
                    old_folder,
                    new_folder,
                    pairing_method,
                    parallel,
                ),
                daemon=True,
            )
        )

        self.worker_thread.start()

    def stop_compare(self):
        if (
            self.worker_thread is None
            or not self.worker_thread.is_alive()
        ):
            return

        self.cancel_event.set()

        self.stop_button.configure(
            state="disabled"
        )

        self.log(
            "Stopping comparison..."
        )

    def parse_progress(
        self,
        message,
    ):
        match = re.search(
            r"\[(\d+)/(\d+)\]\s*(?:✓\s*)?Reading Old:\s*(.+)",
            message,
        )

        if match:
            index = int(
                match.group(1)
            )

            total = int(
                match.group(2)
            )

            filename = (
                match.group(3).strip()
            )

            self.progress_total = total

            self.progress_states[index] = (
                f"[{index}/{total}] "
                f"Reading Old: {filename}"
            )

            self.root.after(
                0,
                self.refresh_progress_log,
            )

            return True

        match = re.search(
            r"\[(\d+)/(\d+)\]\s*(?:✓\s*)?Reading New:\s*(.+)",
            message,
        )

        if match:
            index = int(
                match.group(1)
            )

            total = int(
                match.group(2)
            )

            filename = (
                match.group(3).strip()
            )

            self.progress_total = total

            self.progress_states[index] = (
                f"[{index}/{total}] "
                f"Reading New: {filename}"
            )

            self.root.after(
                0,
                self.refresh_progress_log,
            )

            return True

        match = re.search(
            r"\[(\d+)/(\d+)\]\s*(?:✓\s*)?Comparing:\s*(.+)",
            message,
        )

        if match:
            index = int(
                match.group(1)
            )

            total = int(
                match.group(2)
            )

            filename = (
                match.group(3).strip()
            )

            self.progress_total = total

            self.progress_states[index] = (
                f"[{index}/{total}] "
                f"Comparing: {filename}"
            )

            self.root.after(
                0,
                self.refresh_progress_log,
            )

            return True

        match = re.search(
            r"\[(\d+)/(\d+)\]\s*(?:✓\s*)?Creating output from New Version:\s*(.+)",
            message,
        )

        if match:
            index = int(
                match.group(1)
            )

            total = int(
                match.group(2)
            )

            filename = (
                match.group(3).strip()
            )

            self.progress_total = total

            self.progress_states[index] = (
                f"[{index}/{total}] "
                f"Creating output: {filename}"
            )

            self.root.after(
                0,
                self.refresh_progress_log,
            )

            return True

        match = re.search(
            r"\[(\d+)/(\d+)\]\s*(?:✓\s*)?Completed:\s*(.+?)\s*\((\d+)\s+cell differences,\s*(\d+)\s+added rows,\s*(\d+)\s+missing rows\)",
            message,
        )

        if match:
            index = int(
                match.group(1)
            )

            total = int(
                match.group(2)
            )

            filename = (
                match.group(3).strip()
            )

            cell_changes = int(
                match.group(4)
            )

            added_rows = int(
                match.group(5)
            )

            missing_rows = int(
                match.group(6)
            )

            self.progress_total = total

            self.progress_states[index] = (
                f"[{index}/{total}] ✓ "
                f"{filename} "
                f"({cell_changes} cell differences, "
                f"{added_rows} added rows, "
                f"{missing_rows} missing rows)"
            )

            self.root.after(
                0,
                self.refresh_progress_log,
            )

            return True

        match = re.search(
            r"\[(\d+)/(\d+)\]\s*(?:✓\s*)?Completed:\s*(.+?)\s*\((\d+)\s+differences\)",
            message,
        )

        if match:
            index = int(
                match.group(1)
            )

            total = int(
                match.group(2)
            )

            filename = (
                match.group(3).strip()
            )

            difference_count = int(
                match.group(4)
            )

            self.progress_total = total

            self.progress_states[index] = (
                f"[{index}/{total}] ✓ "
                f"{filename} "
                f"({difference_count} differences)"
            )

            self.root.after(
                0,
                self.refresh_progress_log,
            )

            return True

        return False

    def refresh_progress_log(self):
        current_text = (
            self.log_text.get(
                "1.0",
                "end",
            )
        )

        lines = current_text.splitlines()

        filtered_lines = []

        for line in lines:
            if re.match(
                r"^\[\d+/\d+\]",
                line,
            ):
                continue

            filtered_lines.append(line)

        while (
            filtered_lines
            and not filtered_lines[-1].strip()
        ):
            filtered_lines.pop()

        progress_lines = [
            self.progress_states[index]
            for index in sorted(
                self.progress_states
            )
        ]

        new_lines = (
            filtered_lines
            + progress_lines
        )

        new_text = "\n".join(
            new_lines
        )

        if new_text:
            new_text += "\n"

        self.log_text.configure(
            state="normal"
        )

        self.log_text.delete(
            "1.0",
            "end",
        )

        self.log_text.insert(
            "end",
            new_text,
        )

        self.log_text.see(
            "end"
        )

        self.log_text.configure(
            state="disabled"
        )

    def run_compare(
        self,
        old_folder,
        new_folder,
        pairing_method,
        parallel,
    ):
        try:
            def progress_callback(message):
                handled = (
                    self.parse_progress(
                        message
                    )
                )

                if not handled:
                    self.log(
                        message
                    )

            result = compare_versions(
                old_folder=old_folder,
                new_folder=new_folder,
                pairing_method=pairing_method,
                parallel=parallel,
                progress_callback=(
                    progress_callback
                ),
                cancel_event=(
                    self.cancel_event
                ),
            )

            if result.get(
                "cancelled",
                False,
            ):
                self.root.after(
                    0,
                    self.show_stopped,
                )
                return

            self.root.after(
                0,
                lambda result=result: (
                    self.show_final_result(
                        result
                    )
                ),
            )

        except InterruptedError:
            self.root.after(
                0,
                self.show_stopped,
            )

        except Exception as exc:
            error_message = str(exc)

            self.root.after(
                0,
                lambda message=error_message: (
                    self.show_error(
                        message
                    )
                ),
            )

    def count_rows(
        self,
        value,
    ):
        if value is None:
            return 0

        if isinstance(
            value,
            bool,
        ):
            return int(value)

        if isinstance(
            value,
            (int, float),
        ):
            return int(value)

        if isinstance(
            value,
            dict,
        ):
            return sum(
                self.count_rows(item)
                for item in value.values()
            )

        if isinstance(
            value,
            (list, tuple, set),
        ):
            return len(value)

        return 0

    def get_row_count(
        self,
        item,
        main_key,
        count_key,
    ):
        if item.get(count_key) is not None:
            return self.count_rows(
                item.get(count_key)
            )

        return self.count_rows(
            item.get(
                main_key,
                0,
            )
        )

    def show_final_result(
        self,
        result,
    ):
        results = result.get(
            "results",
            [],
        )

        output_dir = result.get(
            "output_dir"
        )

        unmatched_old = result.get(
            "unmatched_old",
            [],
        )

        unmatched_new = result.get(
            "unmatched_new",
            [],
        )

        completed_files = 0
        total_cell_changes = 0
        total_added_rows = 0
        total_missing_rows = 0
        error_count = 0

        for item in results:
            status = item.get(
                "status",
                "completed",
            )

            if status == "completed":
                completed_files += 1

            elif status == "error":
                error_count += 1

            total_cell_changes += (
                self.count_rows(
                    item.get(
                        "cell_differences",
                        item.get(
                            "cell_difference_count",
                            0,
                        ),
                    )
                )
            )

            total_added_rows += (
                self.get_row_count(
                    item,
                    "added_rows",
                    "added_row_count",
                )
            )

            total_missing_rows += (
                self.get_row_count(
                    item,
                    "missing_rows",
                    "missing_row_count",
                )
            )

        elapsed_seconds = result.get(
            "elapsed_seconds"
        )

        if elapsed_seconds is None:
            if self.compare_start_time is not None:
                elapsed_seconds = (
                    time.perf_counter()
                    - self.compare_start_time
                )
            else:
                elapsed_seconds = 0

        elapsed_time = result.get(
            "elapsed_time"
        )

        if not elapsed_time:
            elapsed_time = (
                self.format_elapsed_time(
                    elapsed_seconds
                )
            )

        total_files = len(results)

        if output_dir is not None:
            self.output_var.set(
                str(output_dir)
            )

        self.write_final_summary(
            completed_files,
            total_files,
            total_cell_changes,
            total_added_rows,
            total_missing_rows,
            error_count,
            elapsed_time,
            output_dir,
            unmatched_old,
            unmatched_new,
        )

        if output_dir is not None:
            try:
                open_output_folder(
                    output_dir
                )
            except Exception as exc:
                self.append_log(
                    f"Could not open output folder: {exc}"
                )

        self.root.after(
            100,
            lambda: messagebox.showinfo(
                "Completed",
                (
                    "Comparison completed.\n\n"
                    f"Files: {completed_files} / "
                    f"Time: {elapsed_time}\n\n"
                ),
            ),
        )

        self.finish_ui()

    def write_final_summary(
        self,
        completed_files,
        total_files,
        total_cell_changes,
        total_added_rows,
        total_missing_rows,
        error_count,
        elapsed_time,
        output_dir,
        unmatched_old,
        unmatched_new,
    ):
        self.root.after(
            0,
            lambda: self.append_final_summary(
                completed_files,
                total_files,
                total_cell_changes,
                total_added_rows,
                total_missing_rows,
                error_count,
                elapsed_time,
                output_dir,
                unmatched_old,
                unmatched_new,
            ),
        )

    def append_final_summary(
        self,
        completed_files,
        total_files,
        total_cell_changes,
        total_added_rows,
        total_missing_rows,
        error_count,
        elapsed_time,
        output_dir,
        unmatched_old,
        unmatched_new,
    ):
        self.append_log("")
        self.append_log(
            "────────────────────────────────"
        )
        self.append_log(
            "Completed"
        )
        self.append_log(
            f"Files: {completed_files} / {total_files}"
        )
        self.append_log(
            f"Cell changes: {total_cell_changes}"
        )
        self.append_log(
            f"Added rows: {total_added_rows}"
        )
        self.append_log(
            f"Missing rows: {total_missing_rows}"
        )
        self.append_log(
            f"Errors: {error_count}"
        )
        self.append_log(
            f"Elapsed: {elapsed_time}"
        )

        if unmatched_old:
            self.append_log("")
            self.append_log(
                f"Unmatched Old Version files: "
                f"{len(unmatched_old)}"
            )

            for file_path in unmatched_old:
                self.append_log(
                    f"  - {file_path.name}"
                )

        if unmatched_new:
            self.append_log("")
            self.append_log(
                f"Unmatched New Version files: "
                f"{len(unmatched_new)}"
            )

            for file_path in unmatched_new:
                self.append_log(
                    f"  - {file_path.name}"
                )

        self.append_log("")
        self.append_log(
            f"Output folder: {output_dir}"
        )
        self.append_log(
            "Comparison completed."
        )

    def show_stopped(self):
        elapsed_seconds = 0

        if self.compare_start_time is not None:
            elapsed_seconds = (
                time.perf_counter()
                - self.compare_start_time
            )

        elapsed_time = (
            self.format_elapsed_time(
                elapsed_seconds
            )
        )

        self.append_log("")
        self.append_log(
            "Comparison stopped by user."
        )
        self.append_log(
            f"Elapsed: {elapsed_time}"
        )

        messagebox.showinfo(
            "Stopped",
            (
                "Comparison was stopped.\n\n"
                f"Elapsed: {elapsed_time}"
            ),
        )

        self.finish_ui()

    def show_error(
        self,
        message,
    ):
        self.append_log(
            f"ERROR: {message}"
        )

        messagebox.showerror(
            "Error",
            message,
        )

        self.finish_ui()

    def finish_ui(self):
        self.compare_button.configure(
            state="normal"
        )

        self.stop_button.configure(
            state="disabled"
        )

        self.worker_thread = None
        self.compare_start_time = None

    def format_elapsed_time(
        self,
        elapsed_seconds,
    ):
        total_seconds = int(
            elapsed_seconds
        )

        hours = (
            total_seconds // 3600
        )

        minutes = (
            total_seconds % 3600
        ) // 60

        seconds = (
            total_seconds % 60
        )

        if hours > 0:
            return (
                f"{hours:02d}:"
                f"{minutes:02d}:"
                f"{seconds:02d}"
            )

        return (
            f"{minutes:02d}:"
            f"{seconds:02d}"
        )


def run_app():
    root = tk.Tk()

    ExcelVersionDiffApp(
        root
    )

    root.mainloop()


if __name__ == "__main__":
    run_app()