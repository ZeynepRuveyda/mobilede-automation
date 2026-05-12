import sys
from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "Template_file.xlsx"

# Sabit KPI sirasi
KPI_ORDER = [
    "New Listings",
    "Used Listings",
    "New Cars Listings ",
    "Used Cars Listings ",
    "New Motorbikes Listings ",
    "Used Motorbikes Listings ",
    "New Trucks Listings ",
    "Used Trucks Listings ",
    "New Motorhomes Listings ",
    "Used Motorhomes Listings ",
]


def main(date_str: str) -> None:
    exports_dir = BASE_DIR / date_str / "exports"
    in_path = exports_dir / f"merged_original_export_{date_str}.xlsx"
    out_path = exports_dir / f"merged_exports_{date_str}.xlsx"

    # 1) Orijinal merged export (ham birlesim)
    df = pd.read_excel(in_path)
    required = {"Website", "Date Pige", "Name Of Detail Pige", "Detail Pige"}
    missing = required - set(df.columns)
    if missing:
        raise SystemExit(f"Missing columns in merged original export: {missing}")

    # 2) Template'ten dealer sirasi
    twb = pd.read_excel(
        TEMPLATE_PATH,
        sheet_name=0,
        engine="openpyxl",
    )
    # Template kolon basliklari: Year, Month, Day, customer_number, Dealer, Type, Category,
    # Vehicle_condition, KPI, Value, Last week Value
    if "Dealer" not in twb.columns or "KPI" not in twb.columns:
        raise SystemExit("Template_file.xlsx beklenen kolonlara sahip degil.")

    # Dealer sirasi: Template'teki görünüm sırası
    dealer_order = (
        twb["Dealer"].dropna().astype(str).str.strip().drop_duplicates().tolist()
    )

    # 3) Export KPI isimlerini Template KPI formatina cevir
    # Ham export genelde 'new cars', 'used truck', 'new' gibi geliyor.
    raw = df["Name Of Detail Pige"].astype(str).str.strip().str.lower()
    kpi_map = {
        "new": "New Listings",
        "used": "Used Listings",
        "new cars": "New Cars Listings ",
        "used cars": "Used Cars Listings ",
        "new motorbikes": "New Motorbikes Listings ",
        "used motorbikes": "Used Motorbikes Listings ",
        "new truck": "New Trucks Listings ",
        "used truck": "Used Trucks Listings ",
        "new trucks": "New Trucks Listings ",
        "used trucks": "Used Trucks Listings ",
        "new caravan": "New Motorhomes Listings ",
        "used caravan": "Used Motorhomes Listings ",
        "new motorhomes": "New Motorhomes Listings ",
        "used motorhomes": "Used Motorhomes Listings ",
    }
    df["KPI_std"] = raw.map(kpi_map)

    # Sadece bilinen KPI'lar
    df = df[df["KPI_std"].isin(KPI_ORDER)].copy()
    df["Website"] = df["Website"].astype(str).str.strip()

    # 4) Dealer + KPI -> Detail Pige (güncel hafta) haritası
    cur_map = (
        df[["Website", "KPI_std", "Detail Pige"]]
        .dropna(subset=["Detail Pige"])
        .dropna(subset=["Website", "KPI_std"])
        .drop_duplicates(subset=["Website", "KPI_std"])
        .set_index(["Website", "KPI_std"])["Detail Pige"]
        .to_dict()
    )

    # Date Pige: dealer bazında ilk görünen tarih
    date_map = (
        df[["Website", "Date Pige"]]
        .dropna(subset=["Website"])
        .drop_duplicates(subset=["Website"])
        .set_index("Website")["Date Pige"]
        .to_dict()
    )

    # 5) Tum dealer'lari Template sirasi ile yaz; export'ta olmayan dealer varsa yine 10 N/A satiri olustur
    rows = []
    for dealer in dealer_order:
        dealer_str = str(dealer).strip()
        date_val = date_map.get(dealer_str, None)
        for kpi in KPI_ORDER:
            cur_val = cur_map.get((dealer_str, kpi), "N/A")
            # pandas NaN case (export'ta varsa ama bos ise)
            if pd.isna(cur_val):
                cur_val = "N/A"
            rows.append(
                {
                    "Website": dealer_str,
                    "Date Pige": date_val,
                    "Name Of Detail Pige": kpi,
                    "Detail Pige": cur_val,
                }
            )

    result = pd.DataFrame(
        rows,
        columns=[
            "Website",
            "Date Pige",
            "Name Of Detail Pige",
            "Detail Pige",
        ],
    )

    counts = result.groupby("Website").size().unique()
    print("Unique row counts per website:", counts)

    exports_dir.mkdir(parents=True, exist_ok=True)
    result.to_excel(out_path, index=False)
    print(
        f"Rewritten merged export ordered by Template dealers/KPI order at: {out_path}"
    )


if __name__ == "__main__":
    date = sys.argv[1] if len(sys.argv) > 1 else "17.03.2026"
    main(date)


