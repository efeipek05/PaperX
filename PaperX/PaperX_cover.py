from pathlib import Path

BASE = Path(__file__).parent
TEMPLATE_PATH = BASE / "cover_template.tex"
OUTPUT_PATH = BASE / "cover.tex"

ASSETS_DIR = BASE / "assets"
LOGO_TR = "assets/logo_tr.png"
LOGO_EN = "assets/logo_en.png"

# ------------------ LaTeX Escape ------------------
def escape_latex(s: str) -> str:
    replacements = {
        "\\": r"\textbackslash{}",
        "&": r"\&",
        "%": r"\%",
        "$": r"\$",
        "#": r"\#",
        "_": r"\_",
        "{": r"\{",
        "}": r"\}",
        "~": r"\textasciitilde{}",
        "^": r"\textasciicircum{}",
    }
    for k, v in replacements.items():
        s = s.replace(k, v)
    return s


# ------------------ Language ------------------
def ask_language() -> str:
    """
    First question is always in English (as requested).
    """
    while True:
        lang = input("What language do you prefer? (tr/en): ").strip().lower()
        if lang in ("tr", "en"):
            return lang
        print("Please type 'tr' or 'en'.")


def t(lang: str, tr: str, en: str) -> str:
    return tr if lang == "tr" else en


def get_labels(lang: str) -> dict:
    if lang == "en":
        return {
            "UNIVERSITY_NAME": "TOBB ETÜ",
            "DEPARTMENT_NAME": "Mechanical Engineering",
            "REPORT_TYPE": "Lab Report",
            "PREPARED_BY": "Prepared by",
            "NAME_SURNAME": "Name Surname",
            "STUDENT_ID": "Student ID",
            "SUBMISSION_DATE_LABEL": "Submission Date",
        }
    return {
        "UNIVERSITY_NAME": "TOBB ETÜ",
        "DEPARTMENT_NAME": "Makine Mühendisliği",
        "REPORT_TYPE": "Deney Raporu",
        "PREPARED_BY": "Hazırlayan(lar)",
        "NAME_SURNAME": "Ad Soyad",
        "STUDENT_ID": "Öğr. No",
        "SUBMISSION_DATE_LABEL": "Teslim Tarihi",
    }


def main():
    # ---------- Template checks ----------
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            t(
                "tr",
                f"Template bulunamadı: {TEMPLATE_PATH}",
                f"Template not found: {TEMPLATE_PATH}",
            )
        )

    lang = ask_language()
    labels = get_labels(lang)

    # ---------- Logo selection ----------
    logo_path = LOGO_EN if lang == "en" else LOGO_TR

    if not (BASE / logo_path).exists():
        raise FileNotFoundError(
            t(
                lang,
                f"Logo dosyası bulunamadı: {BASE / logo_path}\nBeklenen: assets/logo_tr.png ve assets/logo_en.png",
                f"Logo file not found: {BASE / logo_path}\nExpected: assets/logo_tr.png and assets/logo_en.png",
            )
        )

    # ---------- User inputs (bilingual prompts + bilingual examples) ----------
    ders_kodu = input(
        t(
            lang,
            "Ders kodu (örn: MAK316L): ",
            "Course code (e.g., MAK316L): ",
        )
    ).strip()

    deney_adi = input(
        t(
            lang,
            "Deney başlığı (örn: Deney 3 - Kritik Hız): ",
            "Experiment title (e.g., Experiment 3 - Critical Speed): ",
        )
    ).strip()

    teslim_tarihi = input(
        t(
            lang,
            "Teslim tarihi (örn: 5 Şubat 2026): ",
            "Submission date (e.g., 5 February 2026): ",
        )
    ).strip()

    # ---------- Group size ----------
    while True:
        try:
            n = int(
                input(
                    t(
                        lang,
                        "Grup üye sayısı: ",
                        "Number of group members: ",
                    )
                ).strip()
            )
            if n <= 0:
                raise ValueError
            break
        except ValueError:
            print(
                t(
                    lang,
                    "Lütfen 1 veya daha büyük bir sayı gir.",
                    "Please enter a number that is 1 or greater.",
                )
            )

    # ---------- Members ----------
    rows = []
    for i in range(1, n + 1):
        ad = input(
            t(
                lang,
                f"{i}. üye ad soyad: ",
                f"Member {i} full name: ",
            )
        ).strip()
        no = input(
            t(
                lang,
                f"{i}. üye öğrenci no: ",
                f"Member {i} student ID: ",
            )
        ).strip()
        rows.append(f"{escape_latex(ad)} & {escape_latex(no)} \\\\")
    ogrenci_tablo = "\n".join(rows)

    # ---------- Fill template ----------
    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    filled = template

    # Labels
    for k, v in labels.items():
        filled = filled.replace(f"<<{k}>>", escape_latex(v))

    # Logo + variable fields
    filled = (
        filled
        .replace("<<LOGO_PATH>>", escape_latex(logo_path))
        .replace("<<DERS_KODU>>", escape_latex(ders_kodu))
        .replace("<<DENEY_ADI>>", escape_latex(deney_adi))
        .replace("<<TESLIM_TARIHI>>", escape_latex(teslim_tarihi))
        .replace("<<OGRENCI_TABLO>>", ogrenci_tablo)
    )

    OUTPUT_PATH.write_text(filled, encoding="utf-8")

    print(
        t(
            lang,
            f"\n✅ cover.tex üretildi: {OUTPUT_PATH} (lang = {lang}, logo = {logo_path})",
            f"\n✅ cover.tex generated: {OUTPUT_PATH} (lang = {lang}, logo = {logo_path})",
        )
    )


if __name__ == "__main__":
    main()
