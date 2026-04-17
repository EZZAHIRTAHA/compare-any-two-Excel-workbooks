import argparse
from pathlib import Path
import sys
import pandas as pd


# ── terminal colours (no dependencies) ────────────────────────────────────────
class C:
    RESET  = "\033[0m"
    BOLD   = "\033[1m"
    RED    = "\033[31m"
    GREEN  = "\033[32m"
    YELLOW = "\033[33m"
    CYAN   = "\033[36m"
    DIM    = "\033[2m"

def bold(s):   return f"{C.BOLD}{s}{C.RESET}"
def red(s):    return f"{C.RED}{s}{C.RESET}"
def green(s):  return f"{C.GREEN}{s}{C.RESET}"
def yellow(s): return f"{C.YELLOW}{s}{C.RESET}"
def cyan(s):   return f"{C.CYAN}{s}{C.RESET}"
def dim(s):    return f"{C.DIM}{s}{C.RESET}"


# ── helpers ────────────────────────────────────────────────────────────────────
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
    all_columns = df_a.columns.union(df_b.columns)
    return (
        df_a.reindex(columns=all_columns, fill_value=""),
        df_b.reindex(columns=all_columns, fill_value=""),
    )


def compare_sheet(df_a, df_b):
    """Return (only_in_a, only_in_b, modified) DataFrames."""
    only_in_a = df_a.loc[~df_a.index.isin(df_b.index)]
    only_in_b = df_b.loc[~df_b.index.isin(df_a.index)]

    common = df_a.index.intersection(df_b.index)
    ca, cb = align_dataframes(df_a.loc[common], df_b.loc[common])

    try:
        modified = ca.compare(cb, keep_equal=False)
    except ValueError:
        modified = pd.DataFrame()

    return only_in_a, only_in_b, modified


def modified_row_count(modified: pd.DataFrame) -> int:
    if modified.empty:
        return 0
    if modified.index.nlevels > 1:
        return modified.index.get_level_values(0).nunique()
    return len(modified)


# ── pretty-printing ────────────────────────────────────────────────────────────
BAR_WIDTH = 30

def spark_bar(value, max_value, width=BAR_WIDTH, colour_fn=None):
    """A simple ASCII progress bar."""
    filled = int(round(width * value / max_value)) if max_value else 0
    bar = "█" * filled + "░" * (width - filled)
    return colour_fn(bar) if colour_fn else bar


def print_sheet_result(name, file_a_name, file_b_name, only_in_a, only_in_b, modified):
    n_del = len(only_in_a)
    n_add = len(only_in_b)
    n_mod = modified_row_count(modified)

    print()
    print(bold(f"  Sheet: {cyan(name)}"))
    print(f"  {'Deleted rows':<20} {red(str(n_del)):>6}  {dim('(only in ' + file_a_name + ')')}")
    print(f"  {'Added rows':<20} {green(str(n_add)):>6}  {dim('(only in ' + file_b_name + ')')}")
    print(f"  {'Modified rows':<20} {yellow(str(n_mod)):>6}")


def print_summary(results, file_a_name, file_b_name):
    """Print a consolidated summary table + mini bar charts across all sheets."""
    total_del = sum(len(r["only_in_a"]) for r in results)
    total_add = sum(len(r["only_in_b"]) for r in results)
    total_mod = sum(modified_row_count(r["modified"]) for r in results)
    grand_max  = max(total_del, total_add, total_mod, 1)

    print()
    print("─" * 60)
    print(bold("  SUMMARY"))
    print("─" * 60)
    print(f"  Files compared : {cyan(file_a_name)}  →  {cyan(file_b_name)}")
    print(f"  Sheets compared: {len(results)}")
    print()

    # per-sheet table
    col_w = max(len(r["sheet"]) for r in results) + 2
    header = f"  {'Sheet':<{col_w}} {'Del':>6}  {'Add':>6}  {'Mod':>6}"
    print(dim(header))
    print(dim("  " + "─" * (col_w + 26)))
    for r in results:
        nd = len(r["only_in_a"])
        na = len(r["only_in_b"])
        nm = modified_row_count(r["modified"])
        flag = ""
        if nd == 0 and na == 0 and nm == 0:
            flag = dim("  ✓ identical")
        print(f"  {r['sheet']:<{col_w}} {red(str(nd)):>6}  {green(str(na)):>6}  {yellow(str(nm)):>6}{flag}")

    print()
    print(bold("  Totals"))
    row_max = max(total_del, total_add, total_mod, 1)
    print(f"  {red('Deleted'):>12}  {red(str(total_del)):>5}  {spark_bar(total_del, row_max, colour_fn=red)}")
    print(f"  {green('Added'):>12}  {green(str(total_add)):>5}  {spark_bar(total_add, row_max, colour_fn=green)}")
    print(f"  {yellow('Modified'):>12}  {yellow(str(total_mod)):>5}  {spark_bar(total_mod, row_max, colour_fn=yellow)}")
    print("─" * 60)

    unchanged = sum(
        1 for r in results
        if len(r["only_in_a"]) == 0 and len(r["only_in_b"]) == 0
        and modified_row_count(r["modified"]) == 0
    )
    changed = len(results) - unchanged
    print(f"  {green(str(unchanged))} sheet(s) identical  ·  {yellow(str(changed))} sheet(s) with changes")
    print()


# ── main ───────────────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser(description="Compare two Excel files (all matching sheets).")
    ap.add_argument("file_a", type=Path, help="Older / baseline workbook")
    ap.add_argument("file_b", type=Path, help="Newer workbook to compare to")
    ap.add_argument("--sheet", metavar="NAME",
                    help="Compare only this sheet (name). Omit to compare all matching sheets.")
    ap.add_argument("--key", nargs="+", metavar="COL",
                    help="Column(s) that uniquely identify a row")
    ap.add_argument("--out", type=Path,
                    help="Write an Excel file with added/deleted/modified tabs per sheet")
    args = ap.parse_args()

    # ── discover sheets to compare ─────────────────────────────────────────────
    sheets_a = pd.ExcelFile(args.file_a).sheet_names
    sheets_b = pd.ExcelFile(args.file_b).sheet_names

    if args.sheet:
        if args.sheet not in sheets_a:
            sys.exit(f"[error] Sheet '{args.sheet}' not found in {args.file_a.name}")
        if args.sheet not in sheets_b:
            sys.exit(f"[error] Sheet '{args.sheet}' not found in {args.file_b.name}")
        sheets_to_compare = [args.sheet]
    else:
        sheets_to_compare = [s for s in sheets_a if s in sheets_b]
        only_a = [s for s in sheets_a if s not in sheets_b]
        only_b = [s for s in sheets_b if s not in sheets_a]
        if not sheets_to_compare:
            sys.exit("[error] No sheets with matching names found between the two files.")
        if only_a:
            print(yellow(f"  Sheets only in {args.file_a.name}: {only_a}"))
        if only_b:
            print(yellow(f"  Sheets only in {args.file_b.name}: {only_b}"))

    # ── compare each sheet ─────────────────────────────────────────────────────
    results = []
    for sheet in sheets_to_compare:
        df_a = load_sheet(args.file_a, sheet, args.key)
        df_b = load_sheet(args.file_b, sheet, args.key)
        only_in_a, only_in_b, modified = compare_sheet(df_a, df_b)

        results.append({
            "sheet":     sheet,
            "only_in_a": only_in_a,
            "only_in_b": only_in_b,
            "modified":  modified,
        })

        print_sheet_result(
            sheet,
            args.file_a.name, args.file_b.name,
            only_in_a, only_in_b, modified,
        )

    # ── summary ────────────────────────────────────────────────────────────────
    if len(results) > 1:
        print_summary(results, args.file_a.name, args.file_b.name)

    # ── optional Excel output ──────────────────────────────────────────────────
    if args.out:
        with pd.ExcelWriter(args.out, engine="openpyxl") as xl:
            for r in results:
                safe = r["sheet"][:28]  # Excel tab names max 31 chars
                r["only_in_a"].to_excel(xl, sheet_name=f"{safe}_Deleted")
                r["only_in_b"].to_excel(xl, sheet_name=f"{safe}_Added")
                if not r["modified"].empty:
                    r["modified"].to_excel(xl, sheet_name=f"{safe}_Modified")
                else:
                    pd.DataFrame().to_excel(xl, sheet_name=f"{safe}_Modified")
        print(f"Full detail written to {args.out.resolve()}")


if __name__ == "__main__":
    main()
