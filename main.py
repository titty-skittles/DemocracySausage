from __future__ import annotations

import math
import random
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import pandas as pd


# =========================
# Core election data types
# =========================

@dataclass
class Ballot:
    preferences: List[str]
    weight: float = 1.0
    index: int = 0


class ReturningOfficerPrompt:
    """
    Helper for GUI-based returning officer tie-break decisions.
    """

    def __init__(self, parent: tk.Tk):
        self.parent = parent

    def choose(self, tied_candidates: List[str], context: str) -> str:
        prompt = (
            f"{context}\n\n"
            f"The following candidates are still tied:\n"
            f"{', '.join(tied_candidates)}\n\n"
            f"Type the EXACT candidate name to select."
        )

        while True:
            choice = simpledialog.askstring(
                "Returning Officer Decision",
                prompt,
                parent=self.parent
            )

            if choice is None:
                raise RuntimeError("Count cancelled during returning officer tie-break.")

            choice = choice.strip()
            if choice in tied_candidates:
                return choice

            messagebox.showerror(
                "Invalid choice",
                "Please type one of the tied candidate names exactly as shown."
            )


# =========================
# STV counting engine
# =========================

class STVCounter:
    def __init__(
        self,
        seats: int = 2,
        tie_break_fallback: str = "random",
        ro_prompt: Optional[ReturningOfficerPrompt] = None,
        random_seed: Optional[int] = None,
    ):
        """
        tie_break_fallback:
            - 'random'
            - 'returning_officer'
        """
        self.seats = seats
        self.tie_break_fallback = tie_break_fallback
        self.ro_prompt = ro_prompt
        self.random = random.Random(random_seed)

    @staticmethod
    def normalise_ballot(row: pd.Series, candidates: List[str]) -> List[str]:
        ranked = []
        seen_ranks = set()

        for candidate in candidates:
            value = row[candidate]
            if pd.notna(value):
                try:
                    rank = int(value)
                    if rank < 1:
                        continue
                    if rank in seen_ranks:
                        # Duplicate ranking -> informal ballot
                        return []
                    seen_ranks.add(rank)
                    ranked.append((rank, candidate))
                except (ValueError, TypeError):
                    return []

        if not ranked:
            return []

        ranked.sort(key=lambda x: x[0])

        # Optional formal check: rankings should be consecutive from 1
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
        continuing: List[str]
    ) -> Tuple[Dict[str, float], Dict[str, List[Ballot]]]:
        tallies = {c: 0.0 for c in continuing}
        piles = {c: [] for c in continuing}

        for ballot in ballots:
            candidate = self.next_active_preference(ballot, continuing)
            if candidate is not None and ballot.weight > 0:
                tallies[candidate] += ballot.weight
                piles[candidate].append(ballot)

        return tallies, piles

    def resolve_tie(
        self,
        tied_candidates: List[str],
        round_history: List[Dict[str, float]],
        context: str,
    ) -> str:
        """
        Resolve tie by:
        1. Looking backwards through previous round tallies
        2. Falling back to random or returning officer decision
        """
        if len(tied_candidates) == 1:
            return tied_candidates[0]

        # Look backwards through previous round totals
        for previous_tallies in reversed(round_history):
            scores = {
                c: previous_tallies.get(c, 0.0)
                for c in tied_candidates
            }

            unique_scores = sorted(set(scores.values()))
            if len(unique_scores) > 1:
                if "exclude" in context.lower():
                    # For exclusion, exclude the candidate with the LOWEST prior total
                    lowest = min(scores.values())
                    narrowed = [c for c, v in scores.items() if abs(v - lowest) < 1e-9]
                else:
                    # For election order, choose the candidate with the HIGHEST prior total
                    highest = max(scores.values())
                    narrowed = [c for c, v in scores.items() if abs(v - highest) < 1e-9]

                if len(narrowed) == 1:
                    return narrowed[0]

                tied_candidates = narrowed

        # Fallback
        if self.tie_break_fallback == "random":
            return self.random.choice(tied_candidates)

        if self.tie_break_fallback == "returning_officer":
            if not self.ro_prompt:
                raise RuntimeError("Returning officer prompt is not available.")
            return self.ro_prompt.choose(sorted(tied_candidates), context)

        raise ValueError(f"Unknown tie-break fallback: {self.tie_break_fallback}")

    def count_sheet(
        self,
        df: pd.DataFrame,
        sheet_name: str,
        ballot_number_column: Optional[str] = None,
    ) -> Dict:
        columns = list(df.columns)

        if ballot_number_column and ballot_number_column in columns:
            candidates = [c for c in columns if c != ballot_number_column]
        elif columns and str(columns[0]).strip().lower() in {
            "ballot num", "ballot number", "ballot", "id"
        }:
            candidates = columns[1:]
        else:
            candidates = columns

        ballots: List[Ballot] = []
        informal_count = 0

        for _, row in df.iterrows():
            prefs = self.normalise_ballot(row, candidates)
            if prefs:
                ballots.append(Ballot(preferences=prefs))
            else:
                informal_count += 1

        total_formal = len(ballots)
        quota = math.floor(total_formal / (self.seats + 1)) + 1 if total_formal > 0 else 0

        elected: List[str] = []
        excluded: List[str] = []
        continuing = candidates.copy()

        rounds = []
        tally_history: List[Dict[str, float]] = []

        while len(elected) < self.seats and continuing:
            tallies, piles = self.count_votes(ballots, continuing)
            tally_history.append(tallies.copy())

            round_info = {
                "tallies": {k: round(v, 6) for k, v in tallies.items()},
                "elected_this_round": [],
                "excluded_this_round": None,
                "notes": [],
            }

            # Find candidates at or above quota
            at_quota = [c for c in continuing if tallies[c] >= quota]

            if at_quota:
                # If multiple candidates reach quota together, use tie-break on election order if needed
                # Highest tally first, then previous round totals, then fallback
                ordered = []
                remaining = at_quota.copy()

                while remaining:
                    highest = max(tallies[c] for c in remaining)
                    tied = [c for c in remaining if abs(tallies[c] - highest) < 1e-9]

                    if len(tied) == 1:
                        chosen = tied[0]
                    else:
                        chosen = self.resolve_tie(
                            tied,
                            tally_history[:-1],
                            f"Election order tie in {sheet_name}"
                        )
                        round_info["notes"].append(
                            f"Election order tie between {', '.join(sorted(tied))}; resolved in favour of {chosen}."
                        )

                    ordered.append(chosen)
                    remaining.remove(chosen)

                for candidate in ordered:
                    if candidate not in elected and len(elected) < self.seats:
                        elected.append(candidate)
                        round_info["elected_this_round"].append(candidate)

                        total_for_candidate = tallies[candidate]
                        surplus = total_for_candidate - quota

                        if surplus > 1e-9 and total_for_candidate > 0:
                            transfer_value = surplus / total_for_candidate
                            for ballot in piles[candidate]:
                                ballot.weight *= transfer_value
                                ballot.index += 1
                            round_info["notes"].append(
                                f"{candidate} elected with surplus {surplus:.6f}; transferred at value {transfer_value:.6f}."
                            )
                        else:
                            for ballot in piles[candidate]:
                                ballot.weight = 0.0
                            round_info["notes"].append(
                                f"{candidate} elected exactly on quota; no surplus transferred."
                            )

                continuing = [c for c in continuing if c not in round_info["elected_this_round"]]
                rounds.append(round_info)

                if len(continuing) <= self.seats - len(elected):
                    for c in continuing:
                        elected.append(c)
                    rounds.append({
                        "tallies": {},
                        "elected_this_round": continuing.copy(),
                        "excluded_this_round": None,
                        "notes": ["Remaining candidates filled the remaining vacancies."]
                    })
                    break

                continue

            # No one at quota -> exclude lowest
            lowest = min(tallies[c] for c in continuing)
            tied_lowest = [c for c in continuing if abs(tallies[c] - lowest) < 1e-9]

            if len(tied_lowest) == 1:
                to_exclude = tied_lowest[0]
            else:
                to_exclude = self.resolve_tie(
                    tied_lowest,
                    tally_history[:-1],
                    f"Exclusion tie in {sheet_name}"
                )
                round_info["notes"].append(
                    f"Exclusion tie between {', '.join(sorted(tied_lowest))}; resolved by excluding {to_exclude}."
                )

            excluded.append(to_exclude)
            round_info["excluded_this_round"] = to_exclude

            for ballot in piles[to_exclude]:
                ballot.index += 1

            continuing.remove(to_exclude)
            rounds.append(round_info)

            if len(continuing) <= self.seats - len(elected):
                for c in continuing:
                    elected.append(c)
                rounds.append({
                    "tallies": {},
                    "elected_this_round": continuing.copy(),
                    "excluded_this_round": None,
                    "notes": ["Remaining candidates filled the remaining vacancies."]
                })
                break

        return {
            "sheet_name": sheet_name,
            "total_formal_votes": total_formal,
            "informal_votes": informal_count,
            "quota": quota,
            "winners": elected[:self.seats],
            "rounds": rounds,
        }


# =========================
# Workbook runner
# =========================

def count_workbook(
    file_path: str,
    seats: int,
    tie_break_fallback: str,
    ro_prompt: Optional[ReturningOfficerPrompt] = None,
    random_seed: Optional[int] = None,
) -> List[Dict]:
    workbook = pd.read_excel(file_path, sheet_name=None)

    counter = STVCounter(
        seats=seats,
        tie_break_fallback=tie_break_fallback,
        ro_prompt=ro_prompt,
        random_seed=random_seed,
    )

    results = []
    for sheet_name, df in workbook.items():
        result = counter.count_sheet(df, sheet_name)
        results.append(result)

    return results


def format_results(results: List[Dict]) -> str:
    lines = []

    for result in results:
        lines.append(f"=== {result['sheet_name']} ===")
        lines.append(f"Formal votes: {result['total_formal_votes']}")
        lines.append(f"Informal votes: {result['informal_votes']}")
        lines.append(f"Quota: {result['quota']}")
        lines.append("")

        for i, rnd in enumerate(result["rounds"], start=1):
            lines.append(f"Round {i}")
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
        for winner in result["winners"]:
            lines.append(f"  - {winner}")
        lines.append("")
        lines.append("-" * 60)
        lines.append("")

    return "\n".join(lines)


# =========================
# GUI
# =========================

class ElectionApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Democracy Sausage")
        self.root.geometry("900x700")

        self.file_path_var = tk.StringVar()
        self.seats_var = tk.StringVar(value="2")
        self.tie_break_var = tk.StringVar(value="random")
        self.seed_var = tk.StringVar(value="")

        self.build_ui()

    def build_ui(self):
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill="both", expand=True)

        # File row
        file_frame = ttk.LabelFrame(main, text="Workbook", padding=10)
        file_frame.pack(fill="x", pady=(0, 10))

        ttk.Entry(file_frame, textvariable=self.file_path_var).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(file_frame, text="Browse...", command=self.browse_file).pack(side="left")

        # Settings
        settings = ttk.LabelFrame(main, text="Settings", padding=10)
        settings.pack(fill="x", pady=(0, 10))

        ttk.Label(settings, text="Seats to fill per race:").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Entry(settings, textvariable=self.seats_var, width=10).grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(settings, text="Tie-break fallback:").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=4)
        tie_combo = ttk.Combobox(
            settings,
            textvariable=self.tie_break_var,
            values=["random", "returning_officer"],
            state="readonly",
            width=20
        )
        tie_combo.grid(row=1, column=1, sticky="w", pady=4)

        ttk.Label(settings, text="Random seed (optional):").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Entry(settings, textvariable=self.seed_var, width=10).grid(row=2, column=1, sticky="w", pady=4)

        ttk.Label(
            settings,
            text="Tie-breaks always use previous round totals first. "
                 "This setting is only used if that still does not resolve the tie."
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(8, 0))

        # Buttons
        button_row = ttk.Frame(main)
        button_row.pack(fill="x", pady=(0, 10))

        ttk.Button(button_row, text="Count Election", command=self.run_count).pack(side="left")
        ttk.Button(button_row, text="Save Results...", command=self.save_results).pack(side="left", padx=(8, 0))
        ttk.Button(button_row, text="Clear", command=self.clear_output).pack(side="left", padx=(8, 0))

        # Output
        output_frame = ttk.LabelFrame(main, text="Results", padding=10)
        output_frame.pack(fill="both", expand=True)

        self.output = tk.Text(output_frame, wrap="word")
        self.output.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(output_frame, orient="vertical", command=self.output.yview)
        scrollbar.pack(side="right", fill="y")
        self.output.config(yscrollcommand=scrollbar.set)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select election workbook",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path_var.set(file_path)

    def run_count(self):
        file_path = self.file_path_var.get().strip()
        if not file_path:
            messagebox.showerror("Missing file", "Please select an Excel workbook.")
            return

        try:
            seats = int(self.seats_var.get())
            if seats < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid seats", "Seats must be a whole number greater than 0.")
            return

        seed_text = self.seed_var.get().strip()
        random_seed = None
        if seed_text:
            try:
                random_seed = int(seed_text)
            except ValueError:
                messagebox.showerror("Invalid seed", "Random seed must be a whole number.")
                return

        ro_prompt = ReturningOfficerPrompt(self.root)

        try:
            results = count_workbook(
                file_path=file_path,
                seats=seats,
                tie_break_fallback=self.tie_break_var.get(),
                ro_prompt=ro_prompt,
                random_seed=random_seed,
            )
            formatted = format_results(results)
            self.output.delete("1.0", tk.END)
            self.output.insert(tk.END, formatted)
        except Exception as e:
            messagebox.showerror("Count failed", str(e))

    def save_results(self):
        content = self.output.get("1.0", tk.END).strip()
        if not content:
            messagebox.showinfo("Nothing to save", "There are no results to save yet.")
            return

        file_path = filedialog.asksaveasfilename(
            title="Save results",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if not file_path:
            return

        try:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(content)
            messagebox.showinfo("Saved", "Results saved successfully.")
        except Exception as e:
            messagebox.showerror("Save failed", str(e))

    def clear_output(self):
        self.output.delete("1.0", tk.END)


def main():
    root = tk.Tk()
    app = ElectionApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()