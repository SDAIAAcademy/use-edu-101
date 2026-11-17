# /// script
# requires-python = ">=3.11"
# dependencies = ["xlsxwriter>=3.2"]
# ///
"""Generate the Al-Sindiyan University file pack used by every lab (labs/files/).

    uv run labs/_tools/prepare_files.py

Deterministic: the same seed gives the same files, so every participant and every answer key match.
Everything is synthetic except the open-data table of week 3, which is a committed snapshot of three
World Bank Open Data indicators for 2022 (labs/_tools/public/worldbank_2022.csv, CC BY 4.0).
All names, numbers and organisations are fictional; identity and mobile numbers only look real.
The texts of the files live in texts.py.
"""
from __future__ import annotations

import csv
import random
import shutil
import sys
from pathlib import Path

import xlsxwriter

sys.path.insert(0, str(Path(__file__).parent))
import texts as T  # noqa: E402

TOOLS = Path(__file__).resolve().parent
LABS = TOOLS.parent
OUT = LABS / "files"
REF = LABS / "solutions" / "reference"
SEED = 1449
FIXED = {"created": __import__("datetime").datetime(2026, 11, 1)}


# ------------------------------------------------------------------ helpers
def write_text(path: Path, text: str):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text.strip() + "\n", encoding="utf-8", newline="\n")


def write_csv(path: Path, header, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(header)
        w.writerows(rows)


def book(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(path))
    wb.set_properties({"title": path.stem, "author": "Al-Sindiyan University (fictional)", **FIXED})
    fmt = {
        "head": wb.add_format({"bold": True, "bg_color": "#0F766E", "font_color": "#FFFFFF", "text_wrap": True,
                               "valign": "vcenter", "border": 1}),
        "cell": wb.add_format({"text_wrap": True, "valign": "top", "border": 1}),
        "fill": wb.add_format({"text_wrap": True, "valign": "top", "border": 1, "bg_color": "#FFF7D6"}),
        "num": wb.add_format({"border": 1, "num_format": "0.0"}),
        "int": wb.add_format({"border": 1, "num_format": "#,##0"}),
        "note": wb.add_format({"italic": True, "font_color": "#4B5563", "text_wrap": True}),
        "bold": wb.add_format({"bold": True, "text_wrap": True}),
    }
    return wb, fmt


def sheet(wb, fmt, name, header, rows, widths, fill_cols=(), rtl=True, note=None):
    ws = wb.add_worksheet(name)
    if rtl:
        ws.right_to_left()
    top = 0
    if note:
        ws.merge_range(0, 0, 0, len(header) - 1, note, fmt["note"])
        ws.set_row(0, 34)
        top = 1
    for c, h in enumerate(header):
        ws.write(top, c, h, fmt["head"])
        ws.set_column(c, c, widths[c])
    for r, row in enumerate(rows, start=top + 1):
        for c, v in enumerate(row):
            f = fmt["fill"] if c in fill_cols else fmt["cell"]
            if isinstance(v, float):
                ws.write_number(r, c, v, fmt["num"] if c not in fill_cols else f)
            elif isinstance(v, int):
                ws.write_number(r, c, v, fmt["int"] if c not in fill_cols else f)
            else:
                ws.write(r, c, v, f)
    ws.freeze_panes(top + 1, 0)
    return ws


# ------------------------------------------------------------------ week 1
def week1():
    d = OUT / "week1"
    wb, fmt = book(d / "prediction_log.xlsx")
    rows = [[n, what, prompt, "", "", "", ""] for n, what, prompt in T.TRIALS]
    sheet(wb, fmt, "التجارب", ["#", "ما تفعله", "الأمر", "توقعك قبل التشغيل", "ما حدث", "ما الذي تغيّر", "هل صدق توقعك؟"],
          rows, [4, 34, 52, 30, 30, 30, 14], fill_cols=(3, 4, 5, 6),
          note="اكتب عمود «توقعك» كاملًا قبل أول تشغيل. الخانات الصفراء لك. كل تجربة في محادثة جديدة ما لم يُذكر غير ذلك.")
    ws = wb.add_worksheet("القواعد الثلاث")
    ws.right_to_left()
    ws.set_column(0, 0, 6)
    ws.set_column(1, 1, 90)
    ws.write(0, 0, "#", fmt["head"])
    ws.write(0, 1, "قاعدة عمّا يغيّر مخرج المساعد، بدليل من رقم التجربة", fmt["head"])
    for i in range(3):
        ws.write(i + 1, 0, i + 1, fmt["cell"])
        ws.write(i + 1, 1, "", fmt["fill"])
    wb.close()
    write_text(d / "campus_notice_a.txt", T.NOTICE_A)
    write_text(d / "campus_notice_b.txt", T.NOTICE_B)


# ------------------------------------------------------------------ week 2
def week2():
    d = OUT / "week2"
    wb, fmt = book(d / "scenarios.xlsx")
    rows = [[s["id"], s["college"], s["text"], "", "", ""] for s in T.SCENARIOS]
    sheet(wb, fmt, "السيناريوهات", ["#", "الكلية", "السيناريو", "نوع المهمة", "موضعه على الخريطة", "السبب في سطر"],
          rows, [4, 16, 70, 16, 24, 44], fill_cols=(3, 4, 5),
          note="نوع المهمة: تصنيف، تنبؤ، توصية، توليد، أو «لا شيء: قاعدة مكتوبة». الموضع: أعمق دائرة ينتمي إليها السيناريو.")
    ws = wb.add_worksheet("المواضع")
    ws.right_to_left()
    ws.set_column(0, 0, 30)
    ws.set_column(1, 1, 80)
    ws.write(0, 0, "الموضع", fmt["head"])
    ws.write(0, 1, "متى تختاره", fmt["head"])
    for i, (k, v) in enumerate(T.PLACES, start=1):
        ws.write(i, 0, k, fmt["cell"])
        ws.write(i, 1, v, fmt["cell"])
    wb.close()
    write_text(d / "concept_map_template.svg", T.concept_map_svg())
    write_text(d / "model_system_template.svg", T.model_system_svg())
    write_csv(REF / "scenarios_key.csv", ["id", "task", "place", "model", "system_parts", "high_stakes"],
              [[s["id"], s["task"], s["place"], s["model"], s["system"], s["high"]] for s in T.SCENARIOS])


# ------------------------------------------------------------------ week 3
def load_public():
    with (TOOLS / "public" / "worldbank_2022.csv").open(encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    for r in rows:
        r["net"] = float(r["internet_users_pct"])
        r["life"] = float(r["life_expectancy_years"])
        r["gdp"] = int(r["gdp_per_capita_ppp"])
    return rows


def pearson(xs, ys):
    n = len(xs)
    mx, my = sum(xs) / n, sum(ys) / n
    sxy = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    sxx = sum((x - mx) ** 2 for x in xs)
    syy = sum((y - my) ** 2 for y in ys)
    return sxy / (sxx * syy) ** 0.5


def week3():
    d = OUT / "week3"
    rows = load_public()
    wb, fmt = book(d / "open_data_table.xlsx")
    sheet(wb, fmt, "الجدول", ["الرمز", "الدولة", "Country", "مستخدمو الإنترنت (% من السكان)", "العمر المتوقع عند الولادة (سنة)",
                              "نصيب الفرد من الناتج المحلي (دولار دولي، تعادل القوة الشرائية)"],
          [[r["iso3"], r["country_ar"], r["country_en"], r["net"], r["life"], r["gdp"]] for r in rows],
          [8, 20, 24, 18, 18, 24])
    ws = wb.add_worksheet("المصدر")
    ws.right_to_left()
    ws.set_column(0, 0, 26)
    ws.set_column(1, 1, 100)
    for i, (k, v) in enumerate(T.PUBLIC_SOURCE):
        ws.write(i, 0, k, fmt["bold"])
        ws.write(i, 1, v, fmt["cell"])
    wb.close()
    write_csv(d / "open_data_table.csv", ["iso3", "country_ar", "country_en", "internet_users_pct", "life_expectancy_years",
                                          "gdp_per_capita_ppp"],
              [[r["iso3"], r["country_ar"], r["country_en"], r["net"], r["life"], r["gdp"]] for r in rows])

    by_net = sorted(rows, key=lambda r: -r["net"])
    top5, bottom5 = by_net[:5], by_net[-5:]
    gap_extremes = sum(r["life"] for r in top5) / 5 - sum(r["life"] for r in bottom5) / 5
    hi = [r for r in rows if r["net"] >= 90]
    lo = [r for r in rows if r["net"] < 90]
    facts = {
        "n": len(rows),
        "r_net_life": round(pearson([r["net"] for r in rows], [r["life"] for r in rows]), 2),
        "r_gdp_life": round(pearson([r["gdp"] for r in rows], [r["life"] for r in rows]), 2),
        "r_gdp_net": round(pearson([r["gdp"] for r in rows], [r["net"] for r in rows]), 2),
        "gap_extremes": round(gap_extremes, 1),
        "n_hi": len(hi), "n_lo": len(lo),
        "mean_life_hi": round(sum(r["life"] for r in hi) / len(hi), 1),
        "mean_life_lo": round(sum(r["life"] for r in lo) / len(lo), 1),
    }
    facts["gap_split"] = round(facts["mean_life_hi"] - facts["mean_life_lo"], 1)
    bars = [next(r for r in rows if r["iso3"] == k) for k in ("SAU", "KOR", "GBR", "USA", "DEU")]
    facts["bar_values"] = {r["iso3"]: r["net"] for r in bars}
    facts["bar_height_ratio"] = round((bars[0]["net"] - 90) / (bars[-1]["net"] - 90), 2)
    facts["bar_value_ratio"] = round(bars[0]["net"] / bars[-1]["net"], 2)
    facts["bar_points"] = round(bars[0]["net"] - bars[-1]["net"], 1)
    write_text(d / "chart_1_internet_life.svg", T.chart1_svg(rows, facts))
    write_text(d / "chart_2_internet_bars.svg", T.chart2_svg(bars))

    wb, fmt = book(d / "claim_frame.xlsx")
    for name, claim in (("الرسم 1", T.CLAIM_1.format(gap=facts["gap_extremes"])), ("الرسم 2", T.CLAIM_2)):
        ws = wb.add_worksheet(name)
        ws.right_to_left()
        ws.set_column(0, 0, 30)
        ws.set_column(1, 1, 60)
        ws.set_column(2, 2, 50)
        ws.write(0, 0, "العنصر", fmt["head"])
        ws.write(0, 1, "محاولتك الأولى دون مساعد", fmt["head"])
        ws.write(0, 2, "ما أضفته بعد سؤال المساعد (بلون مختلف)", fmt["head"])
        for i, (k, hint) in enumerate(T.FRAME_ROWS, start=1):
            ws.write(i, 0, k, fmt["bold"])
            ws.write(i, 1, claim if k == "الادعاء كما ورد" else "", fmt["cell"] if k == "الادعاء كما ورد" else fmt["fill"])
            ws.write(i, 2, "", fmt["fill"])
            ws.set_row(i, 60)
            ws.write_comment(i, 0, hint)
    wb.close()
    import json
    write_text(REF / "claim_clinic_facts.json", json.dumps(facts, ensure_ascii=False, indent=2))


# ------------------------------------------------------------------ week 4
FEATURES = ["avg_sentence_words", "sentence_variation", "vocab_variety", "spelling_errors", "common_words_share"]


def essay(rng, kind):
    """One synthetic essay as five measured features. kind: human_l1, human_l2, ai."""
    if kind == "human_l1":
        v = [rng.gauss(19, 4), rng.gauss(8.0, 2.0), rng.gauss(0.62, 0.05), rng.gauss(0.8, 0.6), rng.gauss(0.48, 0.05)]
    elif kind == "human_l2":
        v = [rng.gauss(14, 2.5), rng.gauss(4.4, 1.3), rng.gauss(0.51, 0.045), rng.gauss(0.6, 0.5), rng.gauss(0.56, 0.045)]
    else:
        v = [rng.gauss(21, 2), rng.gauss(4.0, 1.2), rng.gauss(0.55, 0.04), rng.gauss(0.2, 0.2), rng.gauss(0.56, 0.04)]
    v[0] = max(6.0, v[0])
    v[1] = max(0.8, v[1])
    v[2] = min(0.85, max(0.3, v[2]))
    v[3] = max(0.0, v[3])
    v[4] = min(0.8, max(0.3, v[4]))
    return [round(v[0], 1), round(v[1], 1), round(v[2], 3), round(v[3], 1), round(v[4], 3)]


def reference_rule(row):
    """The two-question rule a depth-2 tree reads off the training file (checked in solutions/check.py)."""
    feats = dict(zip(FEATURES, row))
    if feats["sentence_variation"] <= 5.9:
        return "ai"
    return "human"


def week4():
    d = OUT / "week4"
    rng = random.Random(SEED + 4)

    def make(n_h, n_ai, human_kind, prefix):
        rows = []
        kinds = [human_kind] * n_h + ["ai"] * n_ai
        rng.shuffle(kinds)
        for i, k in enumerate(kinds, start=1):
            group = "second_language" if k == "human_l2" else ("first_language" if k == "human_l1" else "not_applicable")
            rows.append([f"{prefix}{i:03d}", *essay(rng, k), group, "ai" if k == "ai" else "human"])
        return rows

    head = ["essay_id", *FEATURES, "writer_group", "written_by"]
    train = make(200, 200, "human_l1", "TR")
    test = make(100, 100, "human_l1", "TE")
    shifted = make(100, 100, "human_l2", "SH")
    write_csv(d / "essays_train.csv", head, train)
    write_csv(d / "essays_test.csv", head, test)
    write_csv(d / "essays_shifted.csv", head, shifted)
    write_text(d / "essays_columns.md", T.ESSAYS_COLUMNS)

    wb, fmt = book(d / "classifier_fallback.xlsx")
    for name, rows in (("اختبار", test), ("بيانات متغيرة", shifted)):
        out = [[*r, reference_rule(r[1:6])] for r in rows]
        sheet(wb, fmt, name, [*head, "tree_says"], out, [10, 12, 12, 12, 12, 12, 16, 12, 12], rtl=False)
    ws = wb.add_worksheet("جدول الأخطاء")
    ws.right_to_left()
    ws.set_column(0, 0, 44)
    ws.set_column(1, 2, 18)
    ws.write(0, 0, "العدّ (اكتب صيغة COUNTIFS في الخانات الصفراء)", fmt["head"])
    ws.write(0, 1, "اختبار", fmt["head"])
    ws.write(0, 2, "بيانات متغيرة", fmt["head"])
    for i, k in enumerate(T.CONFUSION_ROWS, start=1):
        ws.write(i, 0, k, fmt["cell"])
        ws.write(i, 1, "", fmt["fill"])
        ws.write(i, 2, "", fmt["fill"])
    ws.write(len(T.CONFUSION_ROWS) + 2, 0, T.FALLBACK_NOTE, fmt["note"])
    wb.close()


# ------------------------------------------------------------------ week 5
def week5():
    d = OUT / "week5"
    rng = random.Random(SEED + 5)
    rows = []
    #            truth, group, n, flagged
    blocks = [("ذكاء اصطناعي", "لا ينطبق", 500, 460), ("إنسان", "لغته الأولى", 450, 36), ("إنسان", "لغته الثانية", 50, 24)]
    for truth, group, n, flagged in blocks:
        for i in range(n):
            rows.append([truth, group, "نبّه" if i < flagged else "لم ينبّه"])
    rng.shuffle(rows)
    rows = [[f"T{i:04d}", *r] for i, r in enumerate(rows, start=1)]
    wb, fmt = book(d / "detector_trial.xlsx")
    sheet(wb, fmt, "تجربة المورّد", ["رقم النص", "من كتبه فعلًا", "فئة الكاتب", "حكم بصيرة"], rows, [12, 18, 18, 14],
          note="ملف تجربة المورّد كما سلّمه للعمادة: 1,000 نص. اصطناعي لأغراض التدريب.")
    ws = wb.add_worksheet("جدول الأخطاء")
    ws.right_to_left()
    ws.set_column(0, 0, 30)
    ws.set_column(1, 3, 20)
    for c, h in enumerate(["", "بصيرة نبّه", "بصيرة لم ينبّه", "المجموع"]):
        ws.write(0, c, h, fmt["head"])
    for r, h in enumerate(["كتبه ذكاء اصطناعي", "كتبه إنسان", "المجموع"], start=1):
        ws.write(r, 0, h, fmt["bold"])
        for c in range(1, 4):
            ws.write(r, c, "", fmt["fill"])
    for r, h in enumerate(T.TRIAL_QUESTIONS, start=5):
        ws.write(r, 0, h, fmt["cell"])
        ws.write(r, 1, "", fmt["fill"])
    ws = wb.add_worksheet("فصل دراسي")
    ws.right_to_left()
    ws.set_column(0, 0, 60)
    ws.set_column(1, 1, 18)
    ws.write(0, 0, "فصل دراسي في جامعة السنديان", fmt["head"])
    ws.write(0, 1, "العدد", fmt["head"])
    for r, (h, v) in enumerate(T.TERM_ROWS, start=1):
        ws.write(r, 0, h, fmt["cell"])
        if v is None:
            ws.write(r, 1, "", fmt["fill"])
        elif isinstance(v, int):
            ws.write_number(r, 1, v, fmt["int"])
        else:
            ws.write(r, 1, v, fmt["cell"])
    ws = wb.add_worksheet("فئتا الكتّاب")
    ws.right_to_left()
    ws.set_column(0, 0, 30)
    ws.set_column(1, 3, 22)
    for c, h in enumerate(["فئة الكاتب (البشر فقط)", "عدد النصوص", "نبّه عليها بصيرة", "نسبة الإنذار الكاذب"]):
        ws.write(0, c, h, fmt["head"])
    for r, h in enumerate(["لغته الأولى", "لغته الثانية"], start=1):
        ws.write(r, 0, h, fmt["bold"])
        for c in range(1, 4):
            ws.write(r, c, "", fmt["fill"])
    wb.close()
    write_text(d / "vendor_claim.md", T.VENDOR_CLAIM)


# ------------------------------------------------------------------ week 6
def week6():
    d = OUT / "week6"
    wb, fmt = book(d / "coverage_grid.xlsx")
    rows = [[q["id"], q["group"], q["variant"], q["prompt"], "", "", ""] for q in T.COVERAGE]
    sheet(wb, fmt, "شبكة الاختبار", ["#", "المجموعة", "الحالة", "السؤال كما تكتبه", "ملخص الإجابة", "الدرجة (0 إلى 3)", "سبب الدرجة"],
          rows, [5, 12, 20, 70, 40, 12, 36], fill_cols=(4, 5, 6),
          note="كل سؤال في محادثة جديدة. أسئلة المجموعة «لهجات ولغات» تُلصق قبلها حقائق الحرم من campus_facts.txt. قيّم بالمقياس نفسه في كل مرة.")
    ws = wb.add_worksheet("المقياس")
    ws.right_to_left()
    ws.set_column(0, 0, 10)
    ws.set_column(1, 1, 100)
    ws.write(0, 0, "الدرجة", fmt["head"])
    ws.write(0, 1, "متى تُعطى", fmt["head"])
    for i, (k, v) in enumerate(T.SCALE, start=1):
        ws.write(i, 0, k, fmt["cell"])
        ws.write(i, 1, v, fmt["cell"])
    ws = wb.add_worksheet("مفتاح المناطق")
    ws.right_to_left()
    ws.set_column(0, 1, 30)
    ws.write(0, 0, "المنطقة", fmt["head"])
    ws.write(0, 1, "مقر الإمارة", fmt["head"])
    for i, (k, v) in enumerate(T.REGIONS_KEY, start=1):
        ws.write(i, 0, k, fmt["cell"])
        ws.write(i, 1, v, fmt["cell"])
    ws = wb.add_worksheet("المتوسطات")
    ws.right_to_left()
    ws.set_column(0, 0, 24)
    ws.set_column(1, 3, 22)
    for c, h in enumerate(["المجموعة", "متوسط الدرجات", "أضعف حالة", "الفرق بين الأعلى والأدنى"]):
        ws.write(0, c, h, fmt["head"])
    for r, h in enumerate(["أدوار", "مناطق", "لهجات ولغات"], start=1):
        ws.write(r, 0, h, fmt["bold"])
        for c in range(1, 4):
            ws.write(r, c, "", fmt["fill"])
    wb.close()
    write_text(d / "campus_facts.txt", T.CAMPUS_FACTS)


# ------------------------------------------------------------------ week 7
def week7():
    d = OUT / "week7"
    write_text(d / "track_tasks.md", T.track_tasks_md())
    wb, fmt = book(d / "prompt_log.xlsx")
    ws = wb.add_worksheet("المعايير")
    ws.right_to_left()
    ws.set_column(0, 0, 6)
    ws.set_column(1, 1, 60)
    ws.set_column(2, 2, 50)
    for c, h in enumerate(["#", "معيار الجودة (يُكتب قبل أي أمر)", "كيف أعرف أنه تحقق؟"]):
        ws.write(0, c, h, fmt["head"])
    for i in range(1, 5):
        ws.write(i, 0, i, fmt["cell"])
        ws.write(i, 1, "", fmt["fill"])
        ws.write(i, 2, "", fmt["fill"])
    rows = [[r, "", "", "", "", "", "", "", ""] for r in (1, 2, 3)]
    sheet(wb, fmt, "سجل التعديلات", ["الجولة", "الأمر كما أرسلته", "ما غيّرته عن الجولة السابقة ولماذا", "معيار 1 (0 إلى 2)",
                                    "معيار 2 (0 إلى 2)", "معيار 3 (0 إلى 2)", "معيار 4 (0 إلى 2)", "المجموع من 8", "ملاحظات التحقق"],
          rows, [8, 60, 40, 12, 12, 12, 12, 12, 50], fill_cols=(1, 2, 3, 4, 5, 6, 7, 8),
          note="الدرجة 0: لم يتحقق. 1: تحقق جزئيًّا. 2: تحقق. اكتب في ملاحظات التحقق ما راجعته بنفسك في المادة المصدرية وما وجدته.")
    wb.close()


# ------------------------------------------------------------------ week 8
def week8():
    d = OUT / "week8"
    write_text(d / "student_handbook_excerpt.md", T.HANDBOOK)
    write_text(d / "red_team_questions.md", T.RED_TEAM_QUESTIONS)
    write_text(d / "injection_example.txt", T.INJECTION_EXAMPLE)
    wb, fmt = book(d / "verification_checklist.xlsx")
    rows = [[n, q, "", "", "", "", "", ""] for n, q in T.RED_TEAM_SHORT]
    sheet(wb, fmt, "قائمة التحقق", ["#", "السؤال", "الادعاء في الإجابة", "المصدر: أين في الدليل؟", "المنطق: هل يتسق؟",
                                   "المطابقة: الرقم والنص", "عدم اليقين: هل اعترف بما لا يعرف؟", "الحكم: صحيح · خاطئ · لا دليل"],
          rows, [4, 44, 40, 26, 22, 24, 26, 20], fill_cols=(2, 3, 4, 5, 6, 7),
          note="املأ صفًّا لكل إجابة. «لا دليل» حكم صحيح حين لا يذكر الدليل شيئًا عن السؤال.")
    wb.close()
    import json
    write_text(REF / "handbook_facts.json", json.dumps(T.HANDBOOK_FACTS, ensure_ascii=False, indent=2))


# ------------------------------------------------------------------ week 9
def week9():
    d = OUT / "week9"
    write_text(d / "data_rules_sindiyan.md", T.DATA_RULES)
    wb, fmt = book(d / "dataset_scenarios.xlsx")
    rows = [[s["id"], s["text"], "", "", "", ""] for s in T.DATASETS]
    sheet(wb, fmt, "حالات البيانات", ["#", "الحالة", "هل فيها بيانات شخصية أو حساسة؟", "مستوى التصنيف", "هل تدخل المساعد؟ وبأي شرط؟",
                                     "من تسأل قبل المشاركة؟"],
          rows, [4, 80, 28, 18, 36, 26], fill_cols=(2, 3, 4, 5),
          note="مستويات التصنيف وقواعد الجامعة في data_rules_sindiyan.md. اكتب السبب مع كل قرار.")
    wb.close()
    write_text(d / "raw_prompt.txt", T.RAW_PROMPT)
    write_text(d / "tool_task.md", T.TOOL_TASK)
    wb, fmt = book(d / "approval_gate.xlsx")
    rows = [[a["id"], a["text"], "", "", ""] for a in T.ACTIONS]
    sheet(wb, fmt, "بوابة الموافقة", ["#", "ما قد يفعله المساعد حين تُضاف إليه الأدوات", "القرار: مسموح · يحتاج موافقة · محظور",
                                     "هل يُسجَّل؟", "السبب"],
          rows, [4, 70, 30, 14, 44], fill_cols=(2, 3, 4),
          note="أقل الصلاحيات: امنح المساعد ما تحتاجه المهمة فقط. كل ما يغيّر شيئًا خارج المحادثة أو لا يُتراجع عنه يمرّ على إنسان.")
    wb.close()
    import json
    write_csv(REF / "gate_policy.csv", ["id", "decision", "logged"], [[a["id"], a["decision"], a["logged"]] for a in T.ACTIONS])
    write_csv(REF / "datasets_key.csv", ["id", "personal", "level", "may_enter"],
              [[s["id"], s["personal"], s["level"], s["enter"]] for s in T.DATASETS])
    write_text(REF / "identifiers.json", json.dumps(T.IDENTIFIERS, ensure_ascii=False, indent=2))
    write_text(REF / "sanitised_prompt.txt", T.SANITISED_PROMPT)


# ------------------------------------------------------------------ week 10
def week10():
    d = OUT / "week10"
    write_text(d / "use_case_plan_template.md", T.PLAN_TEMPLATE)
    write_text(d / "track_cards.md", T.track_cards_md())
    write_text(d / "peer_review_checklist.md", T.PEER_CHECKLIST)
    write_text(d / "sample_plan_weak.md", T.WEAK_PLAN)
    wb, fmt = book(d / "cost_sheet.xlsx")
    ws = wb.add_worksheet("التكلفة")
    ws.right_to_left()
    ws.set_column(0, 0, 62)
    ws.set_column(1, 2, 18)
    for c, h in enumerate(["البند", "المثال", "مهمتك"]):
        ws.write(0, c, h, fmt["head"])
    for r, (label, value) in enumerate(T.COST_INPUTS, start=1):
        ws.write(r, 0, label, fmt["cell"])
        ws.write_number(r, 1, value, fmt["int"])
        ws.write(r, 2, "", fmt["fill"])
    base = len(T.COST_INPUTS) + 2
    ws.write(base - 1, 0, "النتائج (صيغ: لا تكتب فوقها)", fmt["bold"])
    formulas = [
        ("الوقت في الفصل دون مساعد (دقيقة)", "={c}2*{c}3"),
        ("الوقت في الفصل مع المساعد والتحقق والإعداد (دقيقة)", "={c}6+{c}2*({c}4+{c}5)"),
        ("صافي الوقت الموفَّر في الفصل (دقيقة)", "={c}2*{c}3-({c}6+{c}2*({c}4+{c}5))"),
        ("عدد المرات اللازم ليغطي التوفيرُ وقتَ الإعداد", '=IF({c}3-{c}4-{c}5>0,{c}6/({c}3-{c}4-{c}5),"لا يتحقق")'),
        ("تكلفة الاشتراك في الفصل (ريال)", "={c}7*{c}8"),
    ]
    for i, (label, f) in enumerate(formulas):
        ws.write(base + i, 0, label, fmt["cell"])
        ws.write_formula(base + i, 1, f.format(c="B"), fmt["num"])
        ws.write_formula(base + i, 2, f.format(c="C"), fmt["num"])
    r0 = base + len(formulas) + 1
    ws.write(r0, 0, "تكاليف لا تظهر في الأرقام: اكتب سطرًا عن كل منها في خطتك", fmt["bold"])
    for i, h in enumerate(T.HIDDEN_COSTS, start=1):
        ws.write(r0 + i, 0, h, fmt["cell"])
    wb.close()


def portfolio():
    wb, fmt = book(OUT / "portfolio_tracker.xlsx")
    rows = [[w, a, "", "", "", "", ""] for w, a in T.PORTFOLIO]
    sheet(wb, fmt, "ملف الأعمال", ["الأسبوع", "المخرج الأسبوعي", "تاريخ التسليم", "رقم النسخة", "هل استخدمت الذكاء الاصطناعي؟ في ماذا؟",
                                  "ما تحققت منه بنفسك", "ملاحظة القيود"],
          rows, [10, 54, 16, 12, 40, 40, 36], fill_cols=(2, 3, 4, 5, 6),
          note="ملف أعمالك يمثّل 35% من الدرجة. احتفظ بكل ملف في مجلد يحفظ سجل النسخ، وأرفق سجل محادثات المساعد حين تستخدمه.")
    wb.close()


def main():
    if OUT.exists():
        shutil.rmtree(OUT)
    REF.mkdir(parents=True, exist_ok=True)
    for fn in (week1, week2, week3, week4, week5, week6, week7, week8, week9, week10, portfolio):
        fn()
    n = sum(1 for p in OUT.rglob("*") if p.is_file())
    print(f"wrote {n} files into {OUT.relative_to(LABS.parent)}")


if __name__ == "__main__":
    main()
