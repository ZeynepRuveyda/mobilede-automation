"""
Ham export Excel dosyalarini birlestirir; satirlara/kolonlara hic mudahale etmez.
Cikti: <tarih>/exports/merged_original_export_<tarih>.xlsx
"""
import os
import sys

import pandas as pd

BASE_DIR = os.path.dirname(__file__)


def merge_original_exports(date_str: str) -> str:
    """
    date_str: 'DD.MM.YYYY' (ornek: '17.03.2026')
    exports klasorundeki tum ham .xlsx dosyalarini okur (merged_* dosyalari haric).
    """
    folder = os.path.join(BASE_DIR, date_str, "exports")
    if not os.path.isdir(folder):
        raise FileNotFoundError(f"Klasor bulunamadi: {folder}")

    files = sorted(
        f
        for f in os.listdir(folder)
        if f.lower().endswith((".xlsx", ".xls"))
        and not f.startswith("~$")
        and not f.lower().startswith("merged_exports")
        and not f.lower().startswith("merged_original_export")
    )
    if not files:
        raise FileNotFoundError(f"Ham export dosyasi yok: {folder}")

    frames = []
    for fname in files:
        path = os.path.join(folder, fname)
        df = pd.read_excel(path)
        frames.append(df)

    merged = pd.concat(frames, ignore_index=True)

    out_path = os.path.join(folder, f"merged_original_export_{date_str}.xlsx")
    merged.to_excel(out_path, index=False)
    print(f"Kaydedildi: {out_path} ({len(merged)} satir, {len(files)} dosya)")
    return out_path


if __name__ == "__main__":
    date = sys.argv[1] if len(sys.argv) > 1 else "17.03.2026"
    merge_original_exports(date)
