from __future__ import annotations

import math
import random
from dataclasses import dataclass
from typing import Dict, List, Optional, Protocol, Tuple

import pandas as pd


@dataclass
class Ballot:
    preferences: List[str]
    weight: float = 1.0
    index: int = 0


class TieBreakPrompt(Protocol):
    def choose(self, tied_candidates: List[str], context: str) -> str:
        ...


class STVCounter:
    def __init__(
        self,
        seats: int = 2,
        tie_break_fallback: str = "random",
        ro_prompt: Optional[TieBreakPrompt] = None,
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
                        "action": "No count possible",
                    }
                ],
            }

        while len(elected) < self.seats and continuing:
            tallies, piles = self.count_votes(ballots, continuing)
            tally_history.append(tallies.copy())

            active_ballot_value = sum(tallies.values())

            round_info = {
                "tallies": {candidate: round(votes, 6) for candidate, votes in tallies.items()},
                "elected_this_round": [],
                "excluded_this_round": None,
                "notes": [],
                "action": "",
                "continuing_candidates": continuing[:],
                "active_ballot_value": round(active_ballot_value, 6),
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
                            f"Election order tie between {', '.join(sorted(tied))}. "
                            f"Tie resolved in favour of {chosen}."
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
                            f"{candidate} reached quota and was elected. "
                            f"Surplus = {surplus:.6f}. Transfer value = {transfer_value:.6f}."
                        )
                    else:
                        for ballot in piles[candidate]:
                            ballot.weight = 0.0
                        round_info["notes"].append(
                            f"{candidate} reached quota exactly and was elected. "
                            f"No surplus was transferred."
                        )

                round_info["action"] = f"Elect {', '.join(round_info['elected_this_round'])}"

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
                            "notes": ["The remaining candidates filled the remaining vacancies."],
                            "action": f"Declare remaining candidates elected: {', '.join(continuing)}",
                            "continuing_candidates": continuing[:],
                            "active_ballot_value": 0.0,
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
                    f"Exclusion tie between {', '.join(sorted(tied_lowest))}. "
                    f"Tie resolved by excluding {to_exclude}."
                )

            round_info["excluded_this_round"] = to_exclude
            round_info["action"] = f"Exclude {to_exclude}"

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
                        "notes": ["The remaining candidates filled the remaining vacancies."],
                        "action": f"Declare remaining candidates elected: {', '.join(continuing)}",
                        "continuing_candidates": continuing[:],
                        "active_ballot_value": 0.0,
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