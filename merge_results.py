# merge_results.py
# Merge multiple interviewer result Excel files into one.
#
# Usage:
#   python merge_results.py --inputs outputs/*.xlsx --out merged_results.xlsx

from __future__ import annotations

import argparse
import glob
import os
from datetime import datetime

import pandas as pd


EVAL_SHEET = "Evaluations"


def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--inputs", nargs="+", required=True, help="Input result xlsx paths (supports glob).")
    p.add_argument("--out", required=True, help="Output xlsx path.")
    return p.parse_args()


def expand_globs(items):
    out = []
    for it in items:
        if any(ch in it for ch in ["*", "?", "["]):
            out.extend(glob.glob(it))
        else:
            out.append(it)
    return sorted(set(out))


def main():
    args = parse_args()
    paths = expand_globs(args.inputs)
    if not paths:
        raise SystemExit("No input files found.")

    merged = []
    for pth in paths:
        try:
            df = pd.read_excel(pth, sheet_name=EVAL_SHEET)
            df["source_file"] = os.path.basename(pth)
            merged.append(df)
        except Exception as e:
            print(f"[WARN] skip {pth}: {e}")

    if not merged:
        raise SystemExit("No valid Evaluations sheets were read.")

    out_df = pd.concat(merged, ignore_index=True)
    # Try to sort by candidate then interviewer then timestamp
    for col in ["timestamp", "candidate_id", "interviewer"]:
        if col not in out_df.columns:
            out_df[col] = ""
    out_df = out_df.sort_values(by=["candidate_id", "interviewer", "timestamp"], kind="stable")

    with pd.ExcelWriter(args.out, engine="openpyxl") as writer:
        out_df.to_excel(writer, index=False, sheet_name="MergedEvaluations")
        meta = pd.DataFrame([{
            "merged_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "input_files": ", ".join(os.path.basename(p) for p in paths),
        }])
        meta.to_excel(writer, index=False, sheet_name="Meta")

    print(f"[OK] merged {len(paths)} files -> {args.out} ({len(out_df)} rows)")


if __name__ == "__main__":
    main()
