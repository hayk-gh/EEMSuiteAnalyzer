# Bill Normalizer — How to Use

This program cleans up the Approver and Bill Entry files, keeps only 2025
bills, removes duplicate entries, and gives you a summary of energy and water
use for every building.

## What you need before you start

1. Python installed on your computer (ask IT if you're not sure it's there).
2. The two files you'll be entering: your Approver file and your Bill Entry
   file. They can be named anything — the program will figure out which is
   which by what you type in.

## How to run it

1. Open a terminal (or Command Prompt on Windows) in the folder where
   `normalize_bills.py` and your two files are saved.
2. Type this and press Enter:

   ```
   python normalize_bills.py
   ```

3. It will ask you two questions. Type the file name each time and press
   Enter:

   ```
   Enter the name (or path) of the Approver file: Approver.xlsx
   Enter the name (or path) of the Bill Entry file: Bill_Entry_Buildings.xlsx
   ```

4. The program does the rest and saves a new file called
   `2025_TotalGHGEmissions.xlsx` in the same folder.

## What's in the finished file

It opens with 3 tabs:

- **Raw Data** — both files, exactly as you gave them, just combined into
  one place.
- **Normalized 2025** — the cleaned-up 2025 data, with duplicates fixed.
- **Totals by Building** — how much electricity (kWh), gas (Therms), and
  water each building used in 2025.
