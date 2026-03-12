from __future__ import annotations

import math
import random
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from typing import Dict, List, Optional, Tuple

import pandas as pd


@dataclass
class Ballot:
    preferences: List[str]
    weight: float = 1.0
    index: int = 0


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


class STVCounter:
    def __init__(
        self,
        seats: int = 2,
        tie_break_fallback: str = "random",
        ro_prompt: Optional[ReturningOfficerPrompt] = None,
        random_seed: Optional[int] = None,
    ):
        self.seats = seats
        self.tie_break_fallback = tie_break_fallback
        self.ro_prompt = ro_prompt
        self.random = random.Random(random_seed)

    @staticmethod
    def normalise_ballot(row: pd.Series, candidates: List[str]) -> List[str]:
        ranked: List[Tuple[int, str]] = []
        seen_ranks = set()

        for candidate in candidates:
            value = row[candidate]

            if pd.isna(value):
                continue

            try:
                if isinstance(value, str):
                    value = value.strip()
                    if value == "":
                        continue
                rank = int(float(value))
            except (ValueError, TypeError):
                return []

            if rank < 1:
                return []

            if rank in seen_ranks:
                return []

            seen_ranks.add(rank)
            ranked.append((rank, candidate))

        if not ranked:
            return []

        ranked.sort(key=lambda x: x[0])
        ranks = [rank for rank, _ in ranked]

        if ranks[0] != 1:
            return []

        if ranks != list(range(1, len(ranks) + 1)):
            return []

        return [candidate for _, candidate in ranked]

    @staticmethod
    def next_active_preference(ballot: Ballot, continuing: List[str]) -> Optional[str]:
        while ballot.index < len(ballot.preferences):
            candidate = ballot.preferences[ballot.index]
            if candidate in continuing:
                return candidate
            ballot.index += 1
        return None

    def count_votes(
        self,
        ballots: List[Ballot],
        continuing: List[str],
    ) -> Tuple[Dict[str, float], Dict[str, List[Ballot]]]:
        tallies = {candidate: 0.0 for candidate in continuing}
        piles = {candidate: [] for candidate in continuing}

        for ballot in ballots:
            if ballot.weight <= 0:
                continue

            candidate = self.next_active_preference(ballot, continuing)
            if candidate is not None:
                tallies[candidate] += ballot.weight
                piles[candidate].append(ballot)

        return tallies, piles

    def resolve_tie(
        self,
        tied_candidates: List[str],
        round_history: List[Dict[str, float]],
        context: str,
    ) -> str:
        if len(tied_candidates) == 1:
            return tied_candidates[0]

        current_tied = tied_candidates[:]

        for previous_tallies in reversed(round_history):
            scores = {candidate: previous_tallies.get(candidate, 0.0) for candidate in current_tied}
            unique_scores = sorted(set(scores.values()))

            if len(unique_scores) <= 1:
                continue

            if "exclude" in context.lower():
                target_value = min(scores.values())
            else:
                target_value = max(scores.values())

            narrowed = [
                candidate
                for candidate, value in scores.items()
                if abs(value - target_value) < 1e-9
            ]

            if len(narrowed) == 1:
                return narrowed[0]

            current_tied = narrowed

        if self.tie_break_fallback == "random":
            return self.random.choice(sorted(current_tied))

        if self.tie_break_fallback == "returning_officer":
            if self.ro_prompt is None:
                raise RuntimeError("Returning officer prompt is not available.")
            return self.ro_prompt.choose(sorted(current_tied), context)

        raise ValueError(f"Unknown tie-break fallback: {self.tie_break_fallback}")

    def count_sheet(
        self,
        df: pd.DataFrame,
        sheet_name: str,
        ballot_number_column: Optional[str] = None,
    ) -> Dict:
        columns = [str(col) for col in df.columns]
        df = df.copy()
        df.columns = columns

        if ballot_number_column and ballot_number_column in columns:
            candidates = [col for col in columns if col != ballot_number_column]
        elif columns and columns[0].strip().lower() in {"ballot num", "ballot number", "ballot", "id"}:
            candidates = columns[1:]
        else:
            candidates = columns

        if len(candidates) < 2:
            raise ValueError(
                f"Sheet '{sheet_name}' does not appear to contain candidate columns."
            )

        ballots: List[Ballot] = []
        informal_count = 0

        for _, row in df.iterrows():
            preferences = self.normalise_ballot(row, candidates)
            if preferences:
                ballots.append(Ballot(preferences=preferences))
            else:
                informal_count += 1

        total_formal = len(ballots)
        quota = math.floor(total_formal / (self.seats + 1)) + 1 if total_formal > 0 else 0

        elected: List[str] = []
        excluded: List[str] = []
        continuing = candidates[:]
        rounds: List[Dict] = []
        tally_history: List[Dict[str, float]] = []

        if total_formal == 0:
            return {
                "sheet_name": sheet_name,
                "total_formal_votes": 0,
                "informal_votes": informal_count,
                "quota": 0,
                "winners": [],
                "rounds": [
                    {
                        "tallies": {},
                        "elected_this_round": [],
                        "excluded_this_round": None,
                        "notes": ["No formal votes were found on this sheet."],
                    }
                ],
            }

        while len(elected) < self.seats and continuing:
            tallies, piles = self.count_votes(ballots, continuing)
            tally_history.append(tallies.copy())

            round_info = {
                "tallies": {candidate: round(votes, 6) for candidate, votes in tallies.items()},
                "elected_this_round": [],
                "excluded_this_round": None,
                "notes": [],
            }

            at_quota = [candidate for candidate in continuing if tallies[candidate] >= quota]

            if at_quota:
                ordered_elections: List[str] = []
                remaining = at_quota[:]

                while remaining:
                    highest = max(tallies[candidate] for candidate in remaining)
                    tied = [
                        candidate
                        for candidate in remaining
                        if abs(tallies[candidate] - highest) < 1e-9
                    ]

                    if len(tied) == 1:
                        chosen = tied[0]
                    else:
                        chosen = self.resolve_tie(
                            tied,
                            tally_history[:-1],
                            f"Election order tie in {sheet_name}",
                        )
                        round_info["notes"].append(
                            f"Election order tie between {', '.join(sorted(tied))}; "
                            f"resolved in favour of {chosen}."
                        )

                    ordered_elections.append(chosen)
                    remaining.remove(chosen)

                for candidate in ordered_elections:
                    if candidate in elected or len(elected) >= self.seats:
                        continue

                    elected.append(candidate)
                    round_info["elected_this_round"].append(candidate)

                    candidate_total = tallies[candidate]
                    surplus = candidate_total - quota

                    if surplus > 1e-9 and candidate_total > 0:
                        transfer_value = surplus / candidate_total
                        for ballot in piles[candidate]:
                            ballot.weight *= transfer_value
                            ballot.index += 1
                        round_info["notes"].append(
                            f"{candidate} elected with surplus {surplus:.6f}; "
                            f"transferred at value {transfer_value:.6f}."
                        )
                    else:
                        for ballot in piles[candidate]:
                            ballot.weight = 0.0
                        round_info["notes"].append(
                            f"{candidate} elected exactly on quota; no surplus transferred."
                        )

                continuing = [
                    candidate
                    for candidate in continuing
                    if candidate not in round_info["elected_this_round"]
                ]
                rounds.append(round_info)

                remaining_seats = self.seats - len(elected)
                if remaining_seats > 0 and len(continuing) <= remaining_seats:
                    elected.extend(continuing)
                    rounds.append(
                        {
                            "tallies": {},
                            "elected_this_round": continuing[:],
                            "excluded_this_round": None,
                            "notes": ["Remaining candidates filled the remaining vacancies."],
                        }
                    )
                    break

                continue

            lowest = min(tallies[candidate] for candidate in continuing)
            tied_lowest = [
                candidate
                for candidate in continuing
                if abs(tallies[candidate] - lowest) < 1e-9
            ]

            if len(tied_lowest) == 1:
                to_exclude = tied_lowest[0]
            else:
                to_exclude = self.resolve_tie(
                    tied_lowest,
                    tally_history[:-1],
                    f"Exclusion tie in {sheet_name}",
                )
                round_info["notes"].append(
                    f"Exclusion tie between {', '.join(sorted(tied_lowest))}; "
                    f"resolved by excluding {to_exclude}."
                )

            excluded.append(to_exclude)
            round_info["excluded_this_round"] = to_exclude

            for ballot in piles[to_exclude]:
                ballot.index += 1

            continuing.remove(to_exclude)
            rounds.append(round_info)

            remaining_seats = self.seats - len(elected)
            if remaining_seats > 0 and len(continuing) <= remaining_seats:
                elected.extend(continuing)
                rounds.append(
                    {
                        "tallies": {},
                        "elected_this_round": continuing[:],
                        "excluded_this_round": None,
                        "notes": ["Remaining candidates filled the remaining vacancies."],
                    }
                )
                break

        return {
            "sheet_name": sheet_name,
            "total_formal_votes": total_formal,
            "informal_votes": informal_count,
            "quota": quota,
            "winners": elected[: self.seats],
            "rounds": rounds,
        }


def get_sheet_names(file_path: str) -> List[str]:
    excel_file = pd.ExcelFile(file_path)
    return excel_file.sheet_names


def count_workbook(
    file_path: str,
    seats: int,
    tie_break_fallback: str,
    ro_prompt: Optional[ReturningOfficerPrompt] = None,
    random_seed: Optional[int] = None,
    selected_sheet: Optional[str] = None,
) -> List[Dict]:
    workbook = pd.read_excel(file_path, sheet_name=None)

    counter = STVCounter(
        seats=seats,
        tie_break_fallback=tie_break_fallback,
        ro_prompt=ro_prompt,
        random_seed=random_seed,
    )

    results: List[Dict] = []

    if selected_sheet and selected_sheet != "All sheets":
        if selected_sheet not in workbook:
            raise ValueError(f"Sheet '{selected_sheet}' was not found in the workbook.")
        results.append(counter.count_sheet(workbook[selected_sheet], selected_sheet))
    else:
        for sheet_name, df in workbook.items():
            results.append(counter.count_sheet(df, sheet_name))

    return results


def format_results(results: List[Dict]) -> str:
    lines: List[str] = []

    for result in results:
        lines.append(f"=== {result['sheet_name']} ===")
        lines.append(f"Formal votes: {result['total_formal_votes']}")
        lines.append(f"Informal votes: {result['informal_votes']}")
        lines.append(f"Quota: {result['quota']}")
        lines.append("")

        for round_number, rnd in enumerate(result["rounds"], start=1):
            lines.append(f"Round {round_number}")

            if rnd["tallies"]:
                for candidate, votes in rnd["tallies"].items():
                    lines.append(f"  {candidate}: {votes:.6f}")

            if rnd["elected_this_round"]:
                lines.append(f"  Elected: {', '.join(rnd['elected_this_round'])}")

            if rnd["excluded_this_round"]:
                lines.append(f"  Excluded: {rnd['excluded_this_round']}")

            for note in rnd.get("notes", []):
                lines.append(f"  Note: {note}")

            lines.append("")

        lines.append("Winners:")
        if result["winners"]:
            for winner in result["winners"]:
                lines.append(f"  - {winner}")
        else:
            lines.append("  - No winners determined")
        lines.append("")
        lines.append("-" * 60)
        lines.append("")

    return "\n".join(lines)


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
            formatted = format_results(results)
            self.output.delete("1.0", tk.END)
            self.output.insert(tk.END, formatted)
        except Exception as exc:
            messagebox.showerror("Count failed", str(exc), parent=self.root)

    def save_results(self) -> None:
        content = self.output.get("1.0", tk.END).strip()
        if not content:
            messagebox.showinfo("Nothing to save", "There are no results to save yet.", parent=self.root)
            return

        file_path = filedialog.asksaveasfilename(
            title="Save results",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        )

        if not file_path:
            return

        try:
            with open(file_path, "w", encoding="utf-8") as handle:
                handle.write(content)
            messagebox.showinfo("Saved", "Results saved successfully.", parent=self.root)
        except Exception as exc:
            messagebox.showerror("Save failed", str(exc), parent=self.root)

    def clear_output(self) -> None:
        self.output.delete("1.0", tk.END)


def main() -> None:
    root = tk.Tk()
    ElectionApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()