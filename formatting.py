from __future__ import annotations

from typing import Dict, List


def format_results(results: List[Dict]) -> str:
    lines: List[str] = []

    for result in results:
        lines.append("=" * 72)
        lines.append(f"RACE: {result['sheet_name']}")
        lines.append("=" * 72)
        lines.append(f"Formal votes:   {result['total_formal_votes']}")
        lines.append(f"Informal votes: {result['informal_votes']}")
        lines.append(f"Quota:          {result['quota']}")
        lines.append("")

        for round_number, rnd in enumerate(result["rounds"], start=1):
            lines.append(f"Round #{round_number}")
            lines.append("-" * 72)

            tallies = rnd.get("tallies", {})
            active_ballot_value = sum(tallies.values())

            continuing_candidates = rnd.get("continuing_candidates", [])

            if continuing_candidates:
                lines.append(f"Candidates remaining: {len(continuing_candidates)}")
                lines.append(f"Continuing:           {', '.join(continuing_candidates)}")

            if tallies:
                lines.append(f"Active ballot value:  {active_ballot_value:.6f}")
                if result["quota"] > 0:
                    lines.append(f"Quota:                {result['quota']}")
                lines.append("")
                lines.append("Candidate tallies:")

                ordered = sorted(tallies.items(), key=lambda x: (-x[1], x[0]))
                for candidate, votes in ordered:
                    percent = (votes / active_ballot_value * 100) if active_ballot_value > 0 else 0
                    lines.append(f"  {candidate}: {votes:.6f} ({percent:.2f}%)")

                highest_candidate, highest_votes = ordered[0]
                lowest_candidate, lowest_votes = ordered[-1]

                highest_percent = (highest_votes / active_ballot_value * 100) if active_ballot_value > 0 else 0
                lowest_percent = (lowest_votes / active_ballot_value * 100) if active_ballot_value > 0 else 0

                lines.append("")
                lines.append(
                    f"Highest: {highest_candidate} with {highest_votes:.6f} votes "
                    f"({highest_percent:.2f}%)"
                )
                lines.append(
                    f"Lowest:  {lowest_candidate} with {lowest_votes:.6f} votes "
                    f"({lowest_percent:.2f}%)"
                )

            if rnd.get("action"):
                lines.append("")
                lines.append(f"ACTION: {rnd['action']}")

            if rnd["elected_this_round"]:
                for candidate in rnd["elected_this_round"]:
                    lines.append(f"ELECTED: {candidate}")

            if rnd["excluded_this_round"]:
                lines.append(f"EXCLUDED: {rnd['excluded_this_round']}")

            for note in rnd.get("notes", []):
                lines.append(f"NOTE: {note}")

            lines.append("")

        lines.append("Final Results")
        lines.append("-" * 72)

        if result["winners"]:
            for i, winner in enumerate(result["winners"], start=1):
                lines.append(f"{i}. {winner}")
        else:
            lines.append("No winners determined.")

        lines.append("")
        lines.append("")

    return "\n".join(lines)