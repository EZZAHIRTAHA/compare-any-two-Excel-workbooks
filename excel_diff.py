import argparse
from pathlib import Path
import sys
import pandas as pd


def load_sheet(path: Path, sheet, key_cols):
    """Read an Excel sheet into a DataFrame, set the key as the index, and
    coerce everything else to string so compare() works cleanly."""
    df = pd.read_excel(path, sheet_name=sheet)
    if key_cols:
        try:
            df = df.set_index(key_cols)
        except KeyError as e:
            sys.exit(f"[error] Key column not found in {path}: {e}")
    df = df.sort_index()
    return df.astype(str)


def align_dataframes(df_a, df_b):
    """Align two DataFrames to have the same columns for comparison."""
    all_columns = df_a.columns.union(df_b.columns)
    
    df_a_aligned = df_a.reindex(columns=all_columns, fill_value='')
    df_b_aligned = df_b.reindex(columns=all_columns, fill_value='')
    
    return df_a_aligned, df_b_aligned


def main():
    ap = argparse.ArgumentParser(description="Compare two Excel files.")
    ap.add_argument("file_a", type=Path, help="Older / baseline workbook")
    ap.add_argument("file_b", type=Path, help="Newer workbook to compare to")
    ap.add_argument("--sheet", default=0,
                    help="Sheet name or 0‑based index (default: first sheet)")
    ap.add_argument("--key", nargs="+", metavar="COL",
                    help="Column(s) that uniquely identify a row")
    ap.add_argument("--out", type=Path,
                    help="Write an Excel file with three tabs (added, deleted, modified)")
    args = ap.parse_args()

    df_a = load_sheet(args.file_a, args.sheet, args.key)
    df_b = load_sheet(args.file_b, args.sheet, args.key)

    print(f"Columns in {args.file_a.name}: {list(df_a.columns)}")
    print(f"Columns in {args.file_b.name}: {list(df_b.columns)}")

    only_in_a = df_a.loc[~df_a.index.isin(df_b.index)]
    only_in_b = df_b.loc[~df_b.index.isin(df_a.index)]

    common_indices = df_a.index.intersection(df_b.index)
    common_a = df_a.loc[common_indices]
    common_b = df_b.loc[common_indices]
    
    common_a_aligned, common_b_aligned = align_dataframes(common_a, common_b)
    
    try:
        modified = common_a_aligned.compare(common_b_aligned, keep_equal=False)
    except ValueError as e:
        print(f"Warning: Could not compare DataFrames: {e}")
        modified = pd.DataFrame()

    print(f"Rows only in {args.file_a.name}: {len(only_in_a)}")
    print(f"Rows only in {args.file_b.name}: {len(only_in_b)}")
    
    modified_count = 0
    if not modified.empty:
        if modified.index.nlevels > 1:
            modified_count = modified.index.get_level_values(0).nunique()
        else:
            modified_count = len(modified)
    
    print(f"Rows with modified data: {modified_count}")

    if args.out:
        with pd.ExcelWriter(args.out, engine="openpyxl") as xl:
            only_in_a.to_excel(xl, sheet_name="Deleted_rows")
            only_in_b.to_excel(xl, sheet_name="Added_rows")
            if not modified.empty:
                modified.to_excel(xl, sheet_name="Modified_cells")
            else:
                pd.DataFrame().to_excel(xl, sheet_name="Modified_cells")
        print(f"Full detail written to {args.out.resolve()}")


if __name__ == "__main__":
    main()

