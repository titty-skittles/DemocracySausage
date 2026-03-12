from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from typing import List, Optional

from formatting import format_results, format_official_report
from workbook import count_workbook, get_sheet_names


class ReturningOfficerPrompt:
    def __init__(self, parent: tk.Tk):
        self.parent = parent

    def choose(self, tied_candidates: List[str], context: str) -> str:
        prompt = (
            f"{context}\n\n"
            f"The following candidates are still tied:\n"
            f"{', '.join(tied_candidates)}\n\n"
            f"Type the EXACT candidate name to choose."
        )

        while True:
            choice = simpledialog.askstring(
                "Returning Officer Decision",
                prompt,
                parent=self.parent,
            )

            if choice is None:
                raise RuntimeError("Count cancelled during returning officer tie-break.")

            choice = choice.strip()
            if choice in tied_candidates:
                return choice

            messagebox.showerror(
                "Invalid choice",
                "Please type one of the tied candidate names exactly as shown.",
                parent=self.parent,
            )


class ElectionApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Democracy Sausage")
        self.root.geometry("980x760")

        self.file_path_var = tk.StringVar()
        self.sheet_var = tk.StringVar(value="All sheets")
        self.seats_var = tk.StringVar(value="2")
        self.tie_break_var = tk.StringVar(value="random")
        self.seed_var = tk.StringVar(value="")

        self.sheet_combo: Optional[ttk.Combobox] = None
        self.output: Optional[tk.Text] = None

        self.build_ui()

    def build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill="both", expand=True)

        workbook_frame = ttk.LabelFrame(main, text="Workbook", padding=10)
        workbook_frame.pack(fill="x", pady=(0, 10))

        ttk.Entry(workbook_frame, textvariable=self.file_path_var).pack(
            side="left",
            fill="x",
            expand=True,
            padx=(0, 8),
        )
        ttk.Button(workbook_frame, text="Browse...", command=self.browse_file).pack(side="left")

        settings = ttk.LabelFrame(main, text="Settings", padding=10)
        settings.pack(fill="x", pady=(0, 10))

        ttk.Label(settings, text="Seats to fill per race:").grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=4
        )
        ttk.Entry(settings, textvariable=self.seats_var, width=12).grid(
            row=0, column=1, sticky="w", pady=4
        )

        ttk.Label(settings, text="Race / sheet:").grid(
            row=1, column=0, sticky="w", padx=(0, 10), pady=4
        )
        self.sheet_combo = ttk.Combobox(
            settings,
            textvariable=self.sheet_var,
            values=["All sheets"],
            state="readonly",
            width=32,
        )
        self.sheet_combo.grid(row=1, column=1, sticky="w", pady=4)
        self.sheet_combo.set("All sheets")

        ttk.Label(settings, text="Tie-break fallback:").grid(
            row=2, column=0, sticky="w", padx=(0, 10), pady=4
        )
        tie_combo = ttk.Combobox(
            settings,
            textvariable=self.tie_break_var,
            values=["random", "returning_officer"],
            state="readonly",
            width=20,
        )
        tie_combo.grid(row=2, column=1, sticky="w", pady=4)

        ttk.Label(settings, text="Random seed (optional):").grid(
            row=3, column=0, sticky="w", padx=(0, 10), pady=4
        )
        ttk.Entry(settings, textvariable=self.seed_var, width=12).grid(
            row=3, column=1, sticky="w", pady=4
        )

        ttk.Label(
            settings,
            text=(
                "Tie-breaks always use previous round totals first. "
                "The fallback setting is only used if that does not resolve the tie."
            ),
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(8, 0))

        button_row = ttk.Frame(main)
        button_row.pack(fill="x", pady=(0, 10))

        ttk.Button(button_row, text="Count Election", command=self.run_count).pack(side="left")
        ttk.Button(button_row, text="Save Results...", command=self.save_results).pack(
            side="left", padx=(8, 0)
        )
        ttk.Button(button_row, text="Clear", command=self.clear_output).pack(
            side="left", padx=(8, 0)
        )

        output_frame = ttk.LabelFrame(main, text="Results", padding=10)
        output_frame.pack(fill="both", expand=True)

        self.output = tk.Text(output_frame, wrap="word")
        self.output.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(output_frame, orient="vertical", command=self.output.yview)
        scrollbar.pack(side="right", fill="y")
        self.output.config(yscrollcommand=scrollbar.set)

    def browse_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select election workbook",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )

        if not file_path:
            return

        self.file_path_var.set(file_path)
        self.load_sheet_names(file_path)

    def load_sheet_names(self, file_path: str) -> None:
        if self.sheet_combo is None:
            return

        try:
            sheet_names = get_sheet_names(file_path)
            values = ["All sheets"] + sheet_names
            self.sheet_combo["values"] = values
            self.sheet_combo.set("All sheets")
            self.sheet_var.set("All sheets")
        except Exception as exc:
            self.sheet_combo["values"] = ["All sheets"]
            self.sheet_combo.set("All sheets")
            self.sheet_var.set("All sheets")
            messagebox.showerror(
                "Workbook error",
                f"Could not read sheet names:\n{exc}",
                parent=self.root,
            )

    def run_count(self) -> None:
        file_path = self.file_path_var.get().strip()
        if not file_path:
            messagebox.showerror("Missing file", "Please select an Excel workbook.", parent=self.root)
            return

        if not Path(file_path).exists():
            messagebox.showerror("File not found", "The selected workbook does not exist.", parent=self.root)
            return

        try:
            seats = int(self.seats_var.get().strip())
            if seats < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror(
                "Invalid seats",
                "Seats must be a whole number greater than 0.",
                parent=self.root,
            )
            return

        seed_text = self.seed_var.get().strip()
        random_seed = None
        if seed_text:
            try:
                random_seed = int(seed_text)
            except ValueError:
                messagebox.showerror(
                    "Invalid seed",
                    "Random seed must be a whole number.",
                    parent=self.root,
                )
                return

        ro_prompt = ReturningOfficerPrompt(self.root)

        try:
            results = count_workbook(
                file_path=file_path,
                seats=seats,
                tie_break_fallback=self.tie_break_var.get(),
                ro_prompt=ro_prompt,
                random_seed=random_seed,
                selected_sheet=self.sheet_var.get(),
            )

            self.last_results = results
            formatted = format_results(results)
            self.output.delete("1.0", tk.END)
            self.output.insert(tk.END, formatted)
        except Exception as exc:
            messagebox.showerror("Count failed", str(exc), parent=self.root)

    def save_results(self) -> None:

        if self.output is None:
            return

        content = self.output.get("1.0", tk.END).strip()

        if not content:
            messagebox.showinfo(
                "Nothing to save",
                "There are no results to save yet.",
                parent=self.root,
            )
            return

        file_path = filedialog.asksaveasfilename(
            title="Save results report",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )

        if not file_path:
            return

        try:
            # Use the official report format instead of raw logs
            results = self.last_results
            report = format_official_report(results)

            with open(file_path, "w", encoding="utf-8") as f:
                f.write(report)

            messagebox.showinfo(
                "Saved",
                "Election report saved successfully.",
                parent=self.root
            )

        except Exception as e:
            messagebox.showerror("Save failed", str(e), parent=self.root)

    def clear_output(self) -> None:
        if self.output is not None:
            self.output.delete("1.0", tk.END)