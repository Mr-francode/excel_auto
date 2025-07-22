import argparse
import pandas as pd
import openpyxl

# --- Action Functions ---

def filter_data(df, column, value):
    """Filters the DataFrame based on a column and value."""
    return df[df[column] == value]

def summarize_data(df, group_by_column, agg_column, agg_func):
    """Summarizes the DataFrame by grouping and aggregating."""
    return df.groupby(group_by_column)[agg_column].agg(agg_func).reset_index()

def calculate_column(df, new_column_name, expression):
    """Calculates a new column based on an expression."""
    df[new_column_name] = df.eval(expression)
    return df

def merge_data(df1, df2, on_column, how='inner'):
    """Merges two DataFrames."""
    return pd.merge(df1, df2, on=on_column, how=how)

def sort_data(df, by_columns, ascending=True):
    """Sorts the DataFrame."""
    return df.sort_values(by=by_columns, ascending=ascending)

def rename_columns_data(df, rename_map):
    """Renames columns."""
    return df.rename(columns=rename_map)

def drop_duplicates_data(df, subset=None):
    """Drops duplicate rows."""
    return df.drop_duplicates(subset=subset)

def duplicate_sheet_data(workbook, source_sheet_name, new_sheet_name):
    """Duplicates a sheet in the workbook."""
    source_sheet = workbook[source_sheet_name]
    new_sheet = workbook.copy_worksheet(source_sheet)
    new_sheet.title = new_sheet_name
    return workbook

def update_cells_data(workbook, sheet_name, cell_updates):
    """Updates one or more cells in a specific sheet."""
    sheet = workbook[sheet_name]
    for cell, value in cell_updates.items():
        sheet[cell] = value
    return workbook

# --- Main Application ---

def main():
    parser = argparse.ArgumentParser(
        description='A versatile CLI tool for automating Excel workflows.',
        formatter_class=argparse.RawTextHelpFormatter
    )
    subparsers = parser.add_subparsers(dest='action', required=True, help='The action to perform')

    # --- Filter Action Parser ---
    parser_filter = subparsers.add_parser('filter', help='Filter rows based on a column value')
    parser_filter.add_argument('-i', '--input', required=True, help='Input Excel file')
    parser_filter.add_argument('-o', '--output', required=True, help='Output Excel file')
    parser_filter.add_argument('--column', required=True, help='Column to filter on')
    parser_filter.add_argument('--value', required=True, help='Value to filter for')

    # --- Summarize Action Parser ---
    parser_summarize = subparsers.add_parser('summarize', help='Summarize data by grouping and aggregating')
    parser_summarize.add_argument('-i', '--input', required=True, help='Input Excel file')
    parser_summarize.add_argument('-o', '--output', required=True, help='Output Excel file')
    parser_summarize.add_argument('--group-by', required=True, help='Column to group by')
    parser_summarize.add_argument('--agg-col', required=True, help='Column to aggregate')
    parser_summarize.add_argument('--agg-func', required=True, help='Aggregation function (e.g., mean, sum)')

    # --- Calculate Action Parser ---
    parser_calculate = subparsers.add_parser('calculate', help='Calculate a new column using an expression')
    parser_calculate.add_argument('-i', '--input', required=True, help='Input Excel file')
    parser_calculate.add_argument('-o', '--output', required=True, help='Output Excel file')
    parser_calculate.add_argument('--new-col', required=True, help='Name of the new column')
    parser_calculate.add_argument('--expr', required=True, help='Pandas-compatible expression (e.g., "Salary * 1.1")')

    # --- Merge Action Parser ---
    parser_merge = subparsers.add_parser('merge', help='Merge two Excel files')
    parser_merge.add_argument('--input1', required=True, help='First input Excel file (left)')
    parser_merge.add_argument('--input2', required=True, help='Second input Excel file (right)')
    parser_merge.add_argument('-o', '--output', required=True, help='Output Excel file')
    parser_merge.add_argument('--on', required=True, help='Column to merge on')
    parser_merge.add_argument('--how', default='inner', choices=['inner', 'outer', 'left', 'right'], help='Type of merge')

    # --- Sort Action Parser ---
    parser_sort = subparsers.add_parser('sort', help='Sort rows based on columns')
    parser_sort.add_argument('-i', '--input', required=True, help='Input Excel file')
    parser_sort.add_argument('-o', '--output', required=True, help='Output Excel file')
    parser_sort.add_argument('--by', required=True, nargs='+', help='Column(s) to sort by')
    parser_sort.add_argument('--order', default='asc', choices=['asc', 'desc'], help='Sort order')

    # --- Rename Columns Action Parser ---
    parser_rename = subparsers.add_parser('rename', help='Rename one or more columns')
    parser_rename.add_argument('-i', '--input', required=True, help='Input Excel file')
    parser_rename.add_argument('-o', '--output', required=True, help='Output Excel file')
    parser_rename.add_argument('--map', required=True, help='Mapping of old to new names (e.g., "OldName:NewName,Another:New")')

    # --- Drop Duplicates Action Parser ---
    parser_drop = subparsers.add_parser('drop_duplicates', help='Remove duplicate rows')
    parser_drop.add_argument('-i', '--input', required=True, help='Input Excel file')
    parser_drop.add_argument('-o', '--output', required=True, help='Output Excel file')
    parser_drop.add_argument('--subset', nargs='+', help='Column(s) to consider for identifying duplicates')

    # --- Duplicate Sheet Action Parser ---
    parser_duplicate = subparsers.add_parser('duplicate_sheet', help='Duplicate a sheet in an Excel file')
    parser_duplicate.add_argument('-i', '--input', required=True, help='Input Excel file')
    parser_duplicate.add_argument('-o', '--output', required=True, help='Output Excel file')
    parser_duplicate.add_argument('--source-sheet', required=True, help='Name of the sheet to duplicate')
    parser_duplicate.add_argument('--new-sheet-name', required=True, help='Name for the new duplicated sheet')

    # --- Update Cells Action Parser ---
    parser_update = subparsers.add_parser('update_cells', help='Update one or more cells in a sheet')
    parser_update.add_argument('-i', '--input', required=True, help='Input Excel file')
    parser_update.add_argument('-o', '--output', required=True, help='Output Excel file')
    parser_update.add_argument('--sheet-name', required=True, help='Name of the sheet to update')
    parser_update.add_argument('--updates', required=True, help='Cell updates in the format "A1:NewValue,B2:AnotherValue"')

    args = parser.parse_args()

    # --- Action Dispatch ---
    try:
        # Actions that modify the workbook directly with openpyxl
        if args.action in ['duplicate_sheet', 'update_cells']:
            workbook = openpyxl.load_workbook(args.input)
            if args.action == 'duplicate_sheet':
                workbook = duplicate_sheet_data(workbook, args.source_sheet, args.new_sheet_name)
            elif args.action == 'update_cells':
                cell_updates = dict(item.split(':', 1) for item in args.updates.split(','))
                workbook = update_cells_data(workbook, args.sheet_name, cell_updates)
            workbook.save(args.output)

        # Actions that process data with pandas
        else:
            if args.action == 'merge':
                df1 = pd.read_excel(args.input1)
                df2 = pd.read_excel(args.input2)
            else:
                df = pd.read_excel(args.input)

            if args.action == 'filter':
                result_df = filter_data(df, args.column, args.value)
            elif args.action == 'summarize':
                result_df = summarize_data(df, args.group_by, args.agg_col, args.agg_func)
            elif args.action == 'calculate':
                result_df = calculate_column(df, args.new_col, args.expr)
            elif args.action == 'merge':
                result_df = merge_data(df1, df2, args.on, args.how)
            elif args.action == 'sort':
                result_df = sort_data(df, args.by, ascending=(args.order == 'asc'))
            elif args.action == 'rename':
                rename_map = dict(item.split(':') for item in args.map.split(','))
                result_df = rename_columns_data(df, rename_map)
            elif args.action == 'drop_duplicates':
                result_df = drop_duplicates_data(df, subset=args.subset)
            
            result_df.to_excel(args.output, index=False)

        print(f"Action '{args.action}' completed successfully. Output saved to {args.output}")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == '__main__':
    main()
