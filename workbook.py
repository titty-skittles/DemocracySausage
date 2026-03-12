from __future__ import annotations

from typing import Dict, List, Optional

import pandas as pd

from counter import STVCounter, TieBreakPrompt


def get_sheet_names(file_path: str) -> List[str]:
    excel_file = pd.ExcelFile(file_path)
    return excel_file.sheet_names


def count_workbook(
    file_path: str,
    seats: int,
    tie_break_fallback: str,
    ro_prompt: Optional[TieBreakPrompt] = None,
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