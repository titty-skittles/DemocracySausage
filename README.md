# Democracy Sausage

A simple desktop tool for counting **ranked-choice school elections** using the **Single Transferable Vote (STV)** method.

The program reads ballot data from an Excel spreadsheet and calculates winners using a **Droop quota** and **instant runoff transfers**, supporting elections with **multiple winners per race**.

This tool was designed for **Student Representative Council (SRC) elections**, but can be used for any small ranked-choice election.

---

# Features

- Ranked-choice vote counting (1, 2, 3, 4…)
- Multiple winners per race
- **Single Transferable Vote (STV)** counting
- Droop quota calculation
- Fractional surplus transfers
- Tie-breaking using **previous round totals**
- Optional fallback tie-break methods:
  - Random draw
  - Returning officer decision
- Simple desktop GUI
- Reads **Excel spreadsheets**
- Supports **multiple races in one workbook**

---

# Counting Method

The election is counted using **Single Transferable Vote (STV)**.

## Quota

The quota used is the **Droop quota**:

```
quota = floor(total formal votes / (seats + 1)) + 1
```

Any candidate reaching the quota is elected.

---

## Surplus Transfers

If a candidate exceeds the quota:

- Their **surplus votes** are transferred
- Transfers occur at a **fractional value**

```
transfer value = surplus / total votes for candidate
```

This ensures each ballot contributes proportionally.

---

## Exclusions

If no candidate reaches the quota:

- The **lowest polling candidate is excluded**
- Their ballots transfer to the next valid preference.

---

## Tie-breaking

Ties are resolved using the following order:

1. **Previous round totals**
2. If still tied, the configured fallback method:
   - **Random draw**
   - **Returning officer decision**

---

# Spreadsheet Format

Each **sheet in the workbook represents one race**.

Example:

| Ballot Num | Candidate1 | Candidate2 | Candidate3 | Candidate4 |
|-------------|---------------|------------------|-------------|---------------|
| 1 | 2 | 1 | 3 | 4 |
| 2 | 1 | 4 | 3 | 2 |
| 3 | 3 | 4 | 2 | 1 |

Rules:

- Each row represents **one ballot**
- Numbers represent **preferences**
- `1` = first preference
- Rankings must be **consecutive (1,2,3...)**
- Duplicate rankings are treated as **informal ballots**
- Blank cells mean **no preference**

---

# Installation

## 1. Install Python

Download Python:

https://www.python.org/downloads/

During installation ensure **"Add Python to PATH"** is checked.

---

## 2. Install dependencies

Run:

```
pip install pandas openpyxl
```

---

# Running the Program

Run the GUI:

```
python src_election_gui.py
```

Then:

1. Select your Excel workbook
2. Set the number of seats
3. Choose the tie-break fallback method
4. Click **Count Election**

Results will appear in the results panel.

---

# Example Workflow

1. Export ballot data from Excel
2. Save as:

```
src_election.xlsx
```

3. Open the program
4. Select the file
5. Run the count

The program will produce:

- round-by-round tallies
- exclusions
- surplus transfers
- final winners

---

# Repository Structure

Example structure:

```
Democracy-Sausage/
├─ main.py
└─ README.md
```

---

# Requirements

Python packages:

```
pandas
openpyxl
```

---

# Disclaimer

This software is provided as a **counting tool only**.

Election organisers remain responsible for:

- verifying ballot data
- confirming election rules
- validating results

---

# License

This project is licensed under the **MIT License**.