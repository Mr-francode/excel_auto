# excel_auto

A project dedicated to automating Excel workflows. This repository contains a powerful and extensible command-line tool to streamline repetitive tasks, improve productivity, and enable efficient management of Excel spreadsheets.

## Features

This tool provides a command-line interface to perform the following actions on Excel files:

*   **`filter`**: Filter rows based on a column's value.
*   **`summarize`**: Summarize data by grouping and aggregating.
*   **`calculate`**: Calculate a new column based on a pandas expression.
*   **`merge`**: Combine two Excel files based on a common column.
*   **`sort`**: Sort rows based on one or more columns.
*   **`rename`**: Rename one or more columns.
*   **`drop_duplicates`**: Remove duplicate rows.
*   **`duplicate_sheet`**: Copy a sheet within the same workbook.
*   **`update_cells`**: Change the value of one or more specific cells.

## Usage

The tool is used by specifying an action (e.g., `filter`, `merge`) followed by its specific arguments.

```bash
# General usage
python3 main.py <action> [options]
```

### Actions and Examples

**1. `filter`**

Filters rows based on a specific value in a column.

```bash
python3 main.py filter -i sample_data.xlsx -o filtered.xlsx --column Department --value Sales
```

**2. `summarize`**

Groups data and calculates an aggregate function (e.g., mean, sum, count).

```bash
python3 main.py summarize -i sample_data.xlsx -o summary.xlsx --group-by Department --agg-col Salary --agg-func mean
```

**3. `calculate`**

Adds a new column based on a mathematical expression.

```bash
python3 main.py calculate -i sample_data.xlsx -o with_bonus.xlsx --new-col Bonus --expr "Salary * 0.1"
```

**4. `merge`**

Merges two Excel files. Requires a second input file (`locations.xlsx` in this example).

```bash
# Assumes locations.xlsx has 'Department' and 'Location' columns
python3 main.py merge --input1 sample_data.xlsx --input2 locations.xlsx -o merged_data.xlsx --on Department
```

**5. `sort`**

Sorts the data by one or more columns.

```bash
python3 main.py sort -i sample_data.xlsx -o sorted_data.xlsx --by Salary --order desc
```

**6. `rename`**

Renames columns using a comma-separated map of `OldName:NewName`.

```bash
python3 main.py rename -i sample_data.xlsx -o renamed_data.xlsx --map "Name:Employee Name,Hire_Date:Start Date"
```

**7. `drop_duplicates`**

Removes rows with duplicate values in the specified columns (or all columns if none are specified).

```bash
python3 main.py drop_duplicates -i sample_data.xlsx -o unique_data.xlsx --subset Department
```

**8. `duplicate_sheet`**

Copies an existing sheet to a new sheet within the same workbook.

```bash
python3 main.py duplicate_sheet -i sample_data.xlsx -o duplicated.xlsx --source-sheet Employees --new-sheet-name "Employees (Copy)"
```

**9. `update_cells`**

Updates the value of one or more cells in a specific sheet. The updates are provided as a comma-separated string of `Cell:Value` pairs.

```bash
python3 main.py update_cells -i sample_data.xlsx -o updated.xlsx --sheet-name Employees --updates "A1:Report Title,B1:Status: Final"
```
