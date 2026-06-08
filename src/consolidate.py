"""業種別企業CSVを1ファイルに統合する。

data/raw/ 配下の「大分類_小分類.csv」群を読み込み、
大分類・小分類（ファイル名由来）、社名、地方、都道府県、住所、
代表者名、代表電話を抽出して outputs/ に統合CSV/Excelを出力する。

使い方:
    python src/consolidate.py
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

# src/ を import パスに追加（CCW・ローカル両対応）
sys.path.insert(0, str(Path(__file__).resolve().parent))
from region_map import detect_prefecture, to_region  # noqa: E402

# プロジェクトルート
ROOT = Path(__file__).resolve().parent.parent
RAW_DIR = ROOT / "data" / "raw"
OUT_DIR = ROOT / "outputs"

# 入力CSVの元列名 → 出力列名（住所・地方・都道府県は別途生成）
SOURCE_COLUMNS = {
    "会社名": "社名",
    "住所": "住所",
    "代表者の氏名": "代表者名",
    "電話番号": "代表電話",
    "郵便番号": "郵便番号",
    "法人番号": "法人番号",
}

# 出力列の並び順
OUTPUT_COLUMNS = [
    "大分類", "小分類", "社名", "地方", "都道府県",
    "郵便番号", "住所", "代表者名", "代表電話", "法人番号",
]


def parse_category(stem: str) -> tuple[str, str]:
    """ファイル名（拡張子なし）から大分類・小分類を取り出す。

    最初のアンダースコアで分割する。小分類側にアンダースコアが
    含まれていても大分類は先頭区切りまでとする。

    Args:
        stem: 拡張子を除いたファイル名（例: "製造_業務用機械製造"）。

    Returns:
        (大分類, 小分類) のタプル。区切りが無ければ小分類は空文字。
    """
    if "_" in stem:
        major, minor = stem.split("_", 1)
        return major, minor
    return stem, ""


def load_one(path: Path) -> pd.DataFrame:
    """1ファイルを読み込み、必要列を抽出・整形する。

    Args:
        path: CSVファイルパス。

    Returns:
        整形済みDataFrame。
    """
    major, minor = parse_category(path.stem)

    # BOM除去・全列文字列・空欄はそのまま空文字で読む
    df = pd.read_csv(
        path,
        encoding="utf-8-sig",
        dtype=str,
        keep_default_na=False,
    )

    out = pd.DataFrame()
    out["大分類"] = [major] * len(df)
    out["小分類"] = minor

    # 元列を出力列名にマッピング（存在しない列は空文字で補完）
    for src, dst in SOURCE_COLUMNS.items():
        out[dst] = df[src].fillna("").str.strip() if src in df.columns else ""

    # 住所から都道府県・地方を生成
    out["都道府県"] = out["住所"].map(detect_prefecture)
    out["地方"] = out["都道府県"].map(to_region)

    return out[OUTPUT_COLUMNS]


def fill_rate_report(df: pd.DataFrame) -> str:
    """主要列の充足率（非空率）をテキストで返す。

    Args:
        df: 統合後DataFrame。

    Returns:
        充足率レポート文字列。
    """
    total = len(df)
    lines = ["列ごとの充足率（非空 / 全行）:"]
    for col in OUTPUT_COLUMNS:
        filled = int((df[col].astype(str).str.strip() != "").sum())
        pct = (filled / total * 100) if total else 0.0
        lines.append(f"  {col:　<8}: {filled:>6} / {total:>6}  ({pct:5.1f}%)")
    return "\n".join(lines)


def main() -> None:
    """data/raw の全CSVを統合し outputs に出力する。"""
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    csv_files = sorted(RAW_DIR.glob("*.csv"))
    if not csv_files:
        print(f"[ERROR] {RAW_DIR} にCSVがありません。data/raw/ にファイルを置いてください。")
        sys.exit(1)

    print(f"対象ファイル数: {len(csv_files)}")
    frames: list[pd.DataFrame] = []
    for path in csv_files:
        try:
            part = load_one(path)
            frames.append(part)
            print(f"  読込OK: {path.name}  ({len(part)}件)")
        except Exception as exc:  # noqa: BLE001
            print(f"  [WARN] 読込失敗: {path.name} -> {exc}")

    merged = pd.concat(frames, ignore_index=True)

    # 都道府県を判定できなかった行を警告
    no_pref = int((merged["都道府県"].str.strip() == "").sum())
    if no_pref:
        print(f"[NOTE] 都道府県を判定できなかった行: {no_pref}件（住所先頭が都道府県でない等）")

    # 出力
    csv_path = OUT_DIR / "統合_業種別企業リスト.csv"
    xlsx_path = OUT_DIR / "統合_業種別企業リスト.xlsx"
    merged.to_csv(csv_path, index=False, encoding="utf-8-sig")
    try:
        merged.to_excel(xlsx_path, index=False)
        print(f"出力: {xlsx_path.name}")
    except Exception as exc:  # noqa: BLE001
        print(f"[WARN] Excel出力スキップ（openpyxl未導入の可能性）: {exc}")

    print(f"出力: {csv_path.name}")
    print(f"総件数: {len(merged)}")
    print()
    print(fill_rate_report(merged))


if __name__ == "__main__":
    main()
