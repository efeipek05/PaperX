# grafik_uret.py
import os
import re
import json
import shutil
from dataclasses import dataclass
from datetime import datetime

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


# ================== Language (ASK IN ENGLISH FIRST) ==================
def ask_language() -> str:
    while True:
        lang = input("Select language / Dil seç (tr/en): ").strip().lower()
        if lang in ("tr", "en"):
            return lang
        print("Please type 'tr' or 'en'.")


def T(lang: str):
    """Tiny translation helper."""
    tr = {
        "excel_name": "Excel dosyası adı (uzantı yazmana gerek yok): ",
        "n_plots": "Kaç grafik çizmek istiyorsunuz?: ",
        "plot_i": "{i}. grafik için:",
        "degree": "Polinom derecesi kaç olsun? (örn: 2): ",
        "curves": "Kaç eğri çizilecek? (1/2): ",
        "x_col": "X ekseni sütunu (örnek: A): ",
        "y1_col": "1. eğri için Y sütunu (örnek: B): ",
        "y2_col": "2. eğri için Y sütunu (örnek: C): ",
        "bad_int": "❌ Lütfen geçerli bir sayı girin.",
        "bad_degree": "❌ Derece en az 1 olmalı.",
        "bad_curves": "❌ Sadece 1 veya 2 girin.",
        "bad_n": "❌ Grafik sayısı 1 veya daha büyük olmalı.",
        "no_numeric_50": "İlk 50 satırda sayısal veri bulunamadı.",
        "no_numeric_cols": "Seçilen sütunlarda sayısal veri bulunamadı.",
        "insufficient_points": "Polinom fit için yetersiz veri: {n} nokta var, en az {need} gerekir.",
        "insufficient_unique_x": "Polinom fit için X değerleri yeterince farklı değil (unique X: {u}).",
        "done": "✅ Grafikler üretildi: {k} adet PNG -> assets/plots/",
        "warn_exists": "ℹ️ assets/plots klasörü temizlendi (eski grafikler silindi).",
        "meta_written": "✅ plots_meta.json yazıldı.",
        "file_not_found": "❌ Dosya bulunamadı: {p}",
        "slope_label": "Eğim",  # plot üstünde gösterilecek metin
    }
    en = {
        "excel_name": "Excel file name (no need to type .xlsx): ",
        "n_plots": "How many plots do you want to generate?: ",
        "plot_i": "For plot #{i}:",
        "degree": "Polynomial degree? (e.g., 2): ",
        "curves": "How many curves on the same plot? (1/2): ",
        "x_col": "X-axis column (e.g., A): ",
        "y1_col": "Y column for curve 1 (e.g., B): ",
        "y2_col": "Y column for curve 2 (e.g., C): ",
        "bad_int": "❌ Please enter a valid number.",
        "bad_degree": "❌ Degree must be at least 1.",
        "bad_curves": "❌ Please type only 1 or 2.",
        "bad_n": "❌ Number of plots must be >= 1.",
        "no_numeric_50": "No numeric data found in the first 50 rows.",
        "no_numeric_cols": "No numeric data found in the selected columns.",
        "insufficient_points": "Not enough points for polynomial fit: {n} points, need at least {need}.",
        "insufficient_unique_x": "X values are not diverse enough for the polynomial degree (unique X: {u}).",
        "done": "✅ Plots generated: {k} PNG files -> assets/plots/",
        "warn_exists": "ℹ️ Cleaned assets/plots folder (old plots deleted).",
        "meta_written": "✅ plots_meta.json written.",
        "file_not_found": "❌ File not found: {p}",
        "slope_label": "Slope",
    }
    return tr if lang == "tr" else en


# ================== Helpers ==================
def ensure_ext(path_str: str, ext: str) -> str:
    s = (path_str or "").strip().strip('"').strip("'")
    if not s:
        return s
    root, current_ext = os.path.splitext(s)
    if current_ext == "":
        return s + ext
    return s


def col_letter_ok(col: str) -> str:
    col = (col or "").strip().upper()
    if not re.fullmatch(r"[A-Z]{1,3}", col):
        raise ValueError(f"Invalid column letter: {col}")
    return col


def col_letter_to_index(col: str) -> int:
    col = col.strip().upper()
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n - 1


def prepare_plots_folder() -> str:
    plots_dir = os.path.join("assets", "plots")
    if os.path.exists(plots_dir):
        shutil.rmtree(plots_dir)
        os.makedirs(plots_dir, exist_ok=True)
        return plots_dir
    os.makedirs(plots_dir, exist_ok=True)
    return plots_dir


def _first_numeric_row(df: pd.DataFrame, col_indices: list[int], max_rows: int = 50) -> int | None:
    scan = min(max_rows, len(df))
    for r in range(scan):
        ok = True
        for c in col_indices:
            if c >= df.shape[1]:
                ok = False
                break
            v = df.iat[r, c]
            if v is None or (isinstance(v, str) and v.strip() == ""):
                ok = False
                break
            try:
                f = float(v)
                if np.isnan(f):
                    ok = False
                    break
            except Exception:
                ok = False
                break
        if ok:
            return r
    return None


def _safe_str(v) -> str:
    return "" if v is None else str(v).strip()


def read_multi_columns_with_headers(excel_path: str, x_col: str, y_cols: list[str], max_scan_rows: int = 50, msg=None):
    df = pd.read_excel(excel_path, header=None, engine="openpyxl")

    x_idx = col_letter_to_index(x_col)
    y_idx_list = [col_letter_to_index(c) for c in y_cols]
    needed = [x_idx] + y_idx_list

    start_row = _first_numeric_row(df, needed, max_rows=max_scan_rows)
    if start_row is None:
        raise ValueError(msg["no_numeric_50"] if msg else "No numeric data found in first 50 rows.")

    header_row = start_row - 1

    xlabel = _safe_str(df.iat[header_row, x_idx]) if header_row >= 0 and x_idx < df.shape[1] else x_col
    ylabels = []
    for c, idx in zip(y_cols, y_idx_list):
        lab = _safe_str(df.iat[header_row, idx]) if header_row >= 0 and idx < df.shape[1] else c
        ylabels.append(lab)

    xs = []
    ys_list = [[] for _ in y_idx_list]

    def to_float(v):
        try:
            return float(v)
        except Exception:
            return np.nan

    empty_streak = 0
    MAX_EMPTY = 3

    r = start_row
    while r < len(df):
        vals = []
        for idx in needed:
            v = df.iat[r, idx] if idx < df.shape[1] else None
            vals.append(to_float(v))

        all_empty = all(not np.isfinite(v) for v in vals)

        if all_empty:
            empty_streak += 1
            if empty_streak >= MAX_EMPTY:
                break
            r += 1
            continue

        empty_streak = 0

        if any(not np.isfinite(v) for v in vals):
            break

        xs.append(vals[0])
        for j in range(len(y_idx_list)):
            ys_list[j].append(vals[1 + j])
        r += 1

    if len(xs) == 0:
        raise ValueError(msg["no_numeric_cols"] if msg else "No numeric data in selected columns.")

    return xs, ys_list, xlabel, ylabels


def _clean_label(s: str) -> str:
    return (s or "").strip()


def _pick_ylabel(ylabels: list[str]) -> str:
    """
    2 eğride 'Y' yazmasın:
    - Eğer tüm ylabels aynıysa => onu kullan
    - Değilse => ilk dolu olanı kullan; yoksa 'Y'
    """
    cleaned = [_clean_label(x) for x in (ylabels or [])]
    nonempty = [x for x in cleaned if x]
    if not nonempty:
        return "Y"
    uniq = sorted(set(nonempty))
    if len(uniq) == 1:
        return uniq[0]
    # farklı etiketler varsa genel bir ylabel döndür
    return nonempty[0]  # en azından saçma "Y" olmasın


def _legend_labels(ylabels: list[str], fallback_prefix: str = "Y"):
    cleaned = [_clean_label(x) for x in (ylabels or [])]
    # Eğer iki label aynıysa, legend yine gerekir mi?
    # Kullanıcı iki eğriyi ayırt etmek isteyecek -> aynıysa Curve 1/2 gibi ayır.
    nonempty = [x for x in cleaned if x]
    if not nonempty:
        return [f"{fallback_prefix}1", f"{fallback_prefix}2"]
    uniq = sorted(set(nonempty))
    if len(uniq) == 1:
        base = uniq[0]
        return [f"{base} (1)", f"{base} (2)"]
    # farklıysa direkt kullan
    # (sayısı 2'den fazla olursa gene işler)
    out = []
    for i, lab in enumerate(cleaned, start=1):
        out.append(lab if lab else f"{fallback_prefix}{i}")
    return out


def _slope_text_from_poly(coeffs: np.ndarray, x_eval: float, msg: dict) -> str:
    """
    coeffs: np.polyfit sonucu (derece d)
    d=1 ise sabit eğim.
    d>1 ise türev polinomunu alıp x_eval'de eğimi hesapla.
    """
    coeffs = np.asarray(coeffs, dtype=float)
    deg = len(coeffs) - 1
    if deg <= 0:
        return ""
    if deg == 1:
        m = coeffs[0]
        return f"{msg['slope_label']} = {m:.4g}"
    # türev:
    dcoeffs = np.polyder(coeffs)
    m = np.polyval(dcoeffs, x_eval)
    return f"{msg['slope_label']}≈ {m:.4g} (x={x_eval:.4g})"


def make_plot_png(xs, ys_list, degree: int, out_path: str,
                 xlabel: str = "", ylabels: list[str] | None = None, msg=None, lang: str = "tr"):
    xs_arr = np.asarray(xs, dtype=float)
    ylabels = ylabels or [""] * len(ys_list)

    plt.figure()

    # legend isimleri (2 eğri vs.)
    leg_labels = _legend_labels(ylabels) if len(ys_list) > 1 else [(_clean_label(ylabels[0]) or "Y")]

    # eğim yazıları üstte biriktirilecek
    slope_lines = []

    for idx, ys in enumerate(ys_list):
        ys_arr = np.asarray(ys, dtype=float)
        mask = np.isfinite(xs_arr) & np.isfinite(ys_arr)
        x = xs_arr[mask]
        y = ys_arr[mask]

        if len(x) == 0:
            raise ValueError(msg["no_numeric_cols"] if msg else "No numeric data.")

        if len(x) < degree + 1:
            raise ValueError((msg["insufficient_points"] if msg else "Not enough points.")
                             .format(n=len(x), need=degree + 1))

        uniq = np.unique(x).size
        if uniq < degree + 1:
            raise ValueError((msg["insufficient_unique_x"] if msg else "Not enough unique X.")
                             .format(u=uniq))

        xs_line = np.linspace(np.min(x), np.max(x), 400)

        # scatter + fit
        plt.scatter(x, y, label=f"{leg_labels[idx]} data" if len(ys_list) > 1 else None)
        coeffs = np.polyfit(x, y, degree)
        ys_fit = np.polyval(coeffs, xs_line)
        plt.plot(xs_line, ys_fit, label=f"{leg_labels[idx]} fit" if len(ys_list) > 1 else None)

        # eğim hesapla (x ortasında)
        x_mid = float(np.mean([np.min(x), np.max(x)]))
        st = ""
        if degree == 1:
            st = _slope_text_from_poly(coeffs, x_mid, msg or {"slope_label": "Slope"})
        if st:
            # legend label ile birlikte yazalım ki hangi eğri belli olsun
            prefix = leg_labels[idx]
            slope_lines.append(f"{prefix}: {st}")

    # Eksensel etiketler
    plt.xlabel(xlabel if xlabel else "X")

    ylabel_final = _pick_ylabel(ylabels)
    plt.ylabel(ylabel_final)

    # Legend: 2 eğri varsa göster (fit & data karmaşık olmasın diye sadece fitleri göstermek istersen sadeleştiririz)
    if len(ys_list) > 1:
        plt.legend()

    # Eğimleri grafik üstünde göster (sol üst, data kapatmasın diye eksen koordinatıyla)
    if degree == 1 and slope_lines:
        txt = "\n".join(slope_lines)
        plt.gca().text(
            0.02, 0.98, txt,
            transform=plt.gca().transAxes,
            va="top", ha="left",
            bbox=dict(boxstyle="round", alpha=0.2)  # renk belirtmiyorum
        )

    plt.tight_layout()
    plt.savefig(out_path, dpi=200)
    plt.close()


# ================== User Input ==================
@dataclass
class PlotSpec:
    degree: int
    curves: int
    x: str
    y: list[str]


def ask_excel_path(lang: str, msg: dict) -> str:
    p = input(msg["excel_name"]).strip()
    p = ensure_ext(p, ".xlsx")
    p = os.path.abspath(p)
    return p


def ask_plot_specs(lang: str, msg: dict) -> list[PlotSpec]:
    while True:
        try:
            n = int(input(msg["n_plots"]).strip())
            if n <= 0:
                print(msg["bad_n"])
                continue
            break
        except Exception:
            print(msg["bad_int"])

    specs: list[PlotSpec] = []
    for i in range(1, n + 1):
        print(msg["plot_i"].format(i=i))

        while True:
            try:
                degree = int(input(msg["degree"]).strip())
                if degree < 1:
                    print(msg["bad_degree"])
                    continue
                break
            except Exception:
                print(msg["bad_int"])

        while True:
            try:
                curves = int(input(msg["curves"]).strip())
                if curves not in (1, 2):
                    print(msg["bad_curves"])
                    continue
                break
            except Exception:
                print(msg["bad_int"])

        x_col = col_letter_ok(input(msg["x_col"]))
        if curves == 1:
            y1 = col_letter_ok(input(msg["y1_col"]))
            specs.append(PlotSpec(degree=degree, curves=1, x=x_col, y=[y1]))
        else:
            y1 = col_letter_ok(input(msg["y1_col"]))
            y2 = col_letter_ok(input(msg["y2_col"]))
            specs.append(PlotSpec(degree=degree, curves=2, x=x_col, y=[y1, y2]))

    return specs


def main():
    lang = ask_language()
    msg = T(lang)

    excel_path = ask_excel_path(lang, msg)
    if not os.path.exists(excel_path):
        print(msg["file_not_found"].format(p=excel_path))
        return

    specs = ask_plot_specs(lang, msg)

    plots_dir = prepare_plots_folder()
    if os.path.isdir(plots_dir):
        print(msg["warn_exists"])

    meta = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "excel": os.path.basename(excel_path),
        "plots": []
    }

    for i, spec in enumerate(specs, start=1):
        rel = os.path.join("assets", "plots", f"plot{i}.png").replace("\\", "/")
        abs_path = os.path.abspath(rel)

        xs, ys_list, xlabel, ylabels = read_multi_columns_with_headers(
            excel_path, spec.x, spec.y, max_scan_rows=50, msg=msg
        )

        make_plot_png(
            xs, ys_list, spec.degree, abs_path,
            xlabel=xlabel, ylabels=ylabels, msg=msg, lang=lang
        )

        meta["plots"].append({
            "file": f"plot{i}.png",
            "degree": spec.degree,
            "x": spec.x,
            "y": spec.y,
            "xlabel": xlabel,
            "ylabels": ylabels,
            "ylabel_final": _pick_ylabel(ylabels),
        })

    meta_path = os.path.join("assets", "plots", "plots_meta.json")
    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)
    print(msg["meta_written"])

    print(msg["done"].format(k=len(specs)))


if __name__ == "__main__":
    main()
