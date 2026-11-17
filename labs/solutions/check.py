# /// script
# requires-python = ">=3.11"
# dependencies = ["openpyxl>=3.1", "scikit-learn>=1.5"]
# ///
"""Run the reference solution of every lab and check it against the course files.

    uv run labs/solutions/check.py            # all labs
    uv run labs/solutions/check.py 4.1 5.1    # some labs

What runs for real: every computation on the files (the scenario key, the figures of the claim clinic,
a decision tree on the essay files, the detector's error table and base-rate arithmetic, the sanitiser,
the approval gate, the cost sheet, the review of the weak plan).
What runs against a stand-in: every step a participant does in a chat assistant. The stand-in is a
scripted model (standin.json): fixed replies that contain the mistakes a real assistant makes, so each
check shows that the lab's verification step catches them. No model is called and nothing is written.
The no-code builder of lab 4.1 is not run here: a scikit-learn decision tree stands in for it.
"""
from __future__ import annotations

import csv
import json
import re
import sys
from collections import Counter
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
LABS = HERE.parent
FILES = LABS / "files"
REF = HERE / "reference"
STANDIN = json.loads((HERE / "standin.json").read_text(encoding="utf-8"))
RUN: dict[str, callable] = {}


def lab(lab_id):
    def wrap(fn):
        RUN[lab_id] = fn
        return fn
    return wrap


def rows_of(path: Path, sheet: str):
    ws = openpyxl.load_workbook(path, data_only=False)[sheet]
    return [list(r) for r in ws.iter_rows(values_only=True)]


def read_csv(path: Path):
    with path.open(encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def words(s: str) -> int:
    return len(re.findall(r"\S+", s))


def numbers(s: str) -> set[str]:
    return set(re.findall(r"\d[\d,]*(?:\.\d+)?", s))


# ------------------------------------------------------------------ 1.1
@lab("1.1")
def lab_1_1():
    rows = rows_of(FILES / "week1" / "prediction_log.xlsx", "التجارب")[2:]
    assert len(rows) == 8, len(rows)
    s = STANDIN["1.1"]
    same_prompt = rows[0][2] == rows[1][2]
    assert same_prompt and s["t1"] != s["t2"], "the same prompt must give two different sentences"
    n3 = words(s["t3"])
    assert n3 != 10, n3
    a, b = text(FILES / "week1" / "campus_notice_a.txt"), text(FILES / "week1" / "campus_notice_b.txt")
    diff = [(x, y) for x, y in zip(a.split(), b.split()) if x != y]
    assert len(diff) == 2, diff
    assert "خمسة" in s["t5"] and "120" in s["t5"] and "سبعة" in s["t6"] and "80" in s["t6"]
    truth = 4817 * 3926
    said = int(s["t7"].replace(",", ""))
    assert truth == 18911542 and said != truth
    handbook = text(FILES / "week8" / "student_handbook_excerpt.md")
    assert "المادة 47" in s["t8"] and "المادة 47" not in handbook and "الدراجات" not in handbook
    assert "أعتذر" in s["t8_sure"]
    print(f"  trials 1 and 2: same prompt, two different sentences; trial 3: {n3} words where 10 were asked")
    print(f"  trials 5 and 6: the notices differ in {len(diff)} words {diff} and the summaries follow the input")
    print(f"  trial 7: 4817 x 3926 = {truth:,}; the stand-in said {said:,} (off by {abs(truth - said):,})")
    print("  trial 8: an article 47 was invented for a regulation that does not exist, then withdrawn when asked")


# ------------------------------------------------------------------ 2.1
@lab("2.1")
def lab_2_1():
    rows = rows_of(FILES / "week2" / "scenarios.xlsx", "السيناريوهات")[2:]
    key = read_csv(REF / "scenarios_key.csv")
    assert len(rows) == 10 and len(key) == 10
    places = Counter(k["place"] for k in key)
    tasks = Counter(k["task"] for k in key)
    valid = {r[0] for r in rows_of(FILES / "week2" / "scenarios.xlsx", "المواضع")[1:]}
    assert set(places) == valid, set(places) ^ valid
    assert [k["id"] for k in key if k["high_stakes"] == "yes"] == ["1"]
    assert [k["id"] for k in key if k["place"].startswith("خارج")] == ["2", "10"]
    assert next(k for k in key if k["id"] == "8")["place"].startswith("وكيل")
    assert "استرجاع" in next(k for k in key if k["id"] == "7")["system_parts"]
    print("  ten scenarios; places:", dict(places))
    print("  task types:", dict(tasks))
    print("  not AI: 2 and 10; the agent: 8; retrieval around the model: 7; high stakes: 1")


# ------------------------------------------------------------------ 3.1
def pearson(xs, ys):
    n = len(xs)
    mx, my = sum(xs) / n, sum(ys) / n
    return sum((x - mx) * (y - my) for x, y in zip(xs, ys)) / (
        sum((x - mx) ** 2 for x in xs) * sum((y - my) ** 2 for y in ys)) ** 0.5


@lab("3.1")
def lab_3_1():
    rows = rows_of(FILES / "week3" / "open_data_table.xlsx", "الجدول")[1:]
    facts = json.loads(text(REF / "claim_clinic_facts.json"))
    assert len(rows) == 30
    net = {r[0]: r[3] for r in rows}
    life = {r[0]: r[4] for r in rows}
    gdp = {r[0]: r[5] for r in rows}
    order = sorted(net, key=lambda k: -net[k])
    gap = sum(life[k] for k in order[:5]) / 5 - sum(life[k] for k in order[-5:]) / 5
    hi = [k for k in net if net[k] >= 90]
    lo = [k for k in net if net[k] < 90]
    split = round(sum(life[k] for k in hi) / len(hi), 1) - round(sum(life[k] for k in lo) / len(lo), 1)
    r_nl = pearson([net[k] for k in net], [life[k] for k in net])
    r_gl = pearson([gdp[k] for k in net], [life[k] for k in net])
    r_gn = pearson([gdp[k] for k in net], [net[k] for k in net])
    assert round(gap, 1) == facts["gap_extremes"] and round(split, 1) == facts["gap_split"]
    assert round(r_nl, 2) == facts["r_net_life"] and round(r_gl, 2) == facts["r_gdp_life"]
    chart1 = text(FILES / "week3" / "chart_1_internet_life.svg")
    assert str(facts["gap_extremes"]) in chart1
    height_ratio = (net["SAU"] - 90) / (net["DEU"] - 90)
    value_ratio = net["SAU"] / net["DEU"]
    assert round(height_ratio, 2) == facts["bar_height_ratio"] and round(value_ratio, 2) == facts["bar_value_ratio"]
    counter = [k for k in net if net[k] < net["SAU"] and life[k] > life["SAU"]]
    print(f"  chart 1: top five against bottom five = {gap:.1f} years (the headline); split at 90% = {split:.1f} years "
          f"({len(hi)} countries against {len(lo)})")
    print(f"  correlations: internet and life {r_nl:.2f}; income and life {r_gl:.2f}; income and internet {r_gn:.2f}")
    print(f"  {len(counter)} countries have less internet use than Saudi Arabia and a longer life expectancy")
    print(f"  chart 2: Saudi Arabia {net['SAU']} against Germany {net['DEU']}: {net['SAU'] - net['DEU']:.1f} points, "
          f"ratio of values {value_ratio:.2f}, ratio of bar heights {height_ratio:.2f}")


# ------------------------------------------------------------------ 4.1
FEATURES = ["avg_sentence_words", "sentence_variation", "vocab_variety", "spelling_errors", "common_words_share"]


def essays(name):
    rows = read_csv(FILES / "week4" / f"{name}.csv")
    return rows, [[float(r[f]) for f in FEATURES] for r in rows], [r["written_by"] for r in rows]


def table(pred, truth):
    c = Counter(zip(truth, pred))
    return {"tp": c[("ai", "ai")], "fp": c[("human", "ai")], "fn": c[("ai", "human")], "tn": c[("human", "human")],
            "acc": (c[("ai", "ai")] + c[("human", "human")]) / len(truth)}


@lab("4.1")
def lab_4_1():
    from sklearn.tree import DecisionTreeClassifier

    tr_rows, xtr, ytr = essays("essays_train")
    te_rows, xte, yte = essays("essays_test")
    sh_rows, xsh, ysh = essays("essays_shifted")
    assert (len(tr_rows), len(te_rows), len(sh_rows)) == (400, 200, 200)
    assert {r["writer_group"] for r in tr_rows} == {"first_language", "not_applicable"}
    assert {r["writer_group"] for r in sh_rows} == {"second_language", "not_applicable"}
    out = {}
    for depth in (1, 3, None):
        tree = DecisionTreeClassifier(max_depth=depth, random_state=0).fit(xtr, ytr)
        out[depth] = (table(tree.predict(xtr), ytr), table(tree.predict(xte), yte), table(tree.predict(xsh), ysh), tree)
    stump = out[1][3]
    assert FEATURES[stump.tree_.feature[0]] == "sentence_variation" and abs(stump.tree_.threshold[0] - 5.9) < 0.1
    for depth, (a, b, c, _) in out.items():
        print(f"  tree depth {depth}: train {a['acc']:.3f}; test {b['acc']:.3f} (false alarms {b['fp']}, misses {b['fn']}); "
              f"shifted {c['acc']:.3f} (false alarms {c['fp']} of 100 human essays, misses {c['fn']})")
        assert b["acc"] >= 0.85 and c["acc"] <= 0.75 and c["fp"] >= 3 * b["fp"]
    assert out[None][0]["acc"] == 1.0 and out[None][1]["acc"] < 1.0, "the unlimited tree memorises the training file"
    # the fallback workbook carries the one-question rule
    for sheet, truth_rows, ref in (("اختبار", te_rows, out[1][1]), ("بيانات متغيرة", sh_rows, out[1][2])):
        rows = rows_of(FILES / "week4" / "classifier_fallback.xlsx", sheet)[1:]
        t = table([r[8] for r in rows], [r[7] for r in rows])
        assert (t["fp"], t["fn"]) == (ref["fp"], ref["fn"]), (sheet, t, ref)
    print("  the fallback workbook gives the same error table as the depth-1 tree")


# ------------------------------------------------------------------ 5.1
@lab("5.1")
def lab_5_1():
    rows = rows_of(FILES / "week5" / "detector_trial.xlsx", "تجربة المورّد")[2:]
    assert len(rows) == 1000
    c = Counter((r[1], r[3]) for r in rows)
    tp, fn = c[("ذكاء اصطناعي", "نبّه")], c[("ذكاء اصطناعي", "لم ينبّه")]
    fp, tn = c[("إنسان", "نبّه")], c[("إنسان", "لم ينبّه")]
    acc = (tp + tn) / 1000
    assert (tp, fn, fp, tn) == (460, 40, 60, 440) and acc == 0.9
    sens, fpr = tp / (tp + fn), fp / (fp + tn)
    n, base = 5000, 0.04
    ai = round(n * base)
    human = n - ai
    true_alarms, false_alarms = round(ai * sens), round(human * fpr)
    ppv = true_alarms / (true_alarms + false_alarms)
    assert (ai, human, true_alarms, false_alarms) == (200, 4800, 184, 576)
    g = Counter((r[2], r[3]) for r in rows if r[1] == "إنسان")
    l1 = g[("لغته الأولى", "نبّه")] / (g[("لغته الأولى", "نبّه")] + g[("لغته الأولى", "لم ينبّه")])
    l2 = g[("لغته الثانية", "نبّه")] / (g[("لغته الثانية", "نبّه")] + g[("لغته الثانية", "لم ينبّه")])
    assert (round(l1, 2), round(l2, 2)) == (0.08, 0.48)
    print(f"  trial: TP {tp}, FN {fn}, FP {fp}, TN {tn}; accuracy {acc:.0%}; catches {sens:.0%} of AI texts; "
          f"flags {fpr:.0%} of human texts; half of the trial is AI text")
    print(f"  a term of {n:,} at a base rate of {base:.0%}: {true_alarms} true alarms, {false_alarms} false alarms, "
          f"{true_alarms + false_alarms} in all; {ppv:.1%} of alarms are true; {ai - true_alarms} missed")
    for b in (0.01, 0.20):
        t, f = n * b * sens, n * (1 - b) * fpr
        print(f"    base rate {b:.0%}: {t / (t + f):.1%} of alarms are true")
    print(f"  false alarms on human writers: first language {l1:.0%} (of 450), second language {l2:.0%} (of 50)")
    mixed = human * (0.7 * l1 + 0.3 * l2)
    print(f"  if 30% of the term's human writers wrote in a second language: {mixed:.0f} false alarms")


# ------------------------------------------------------------------ 6.1
@lab("6.1")
def lab_6_1():
    grid = rows_of(FILES / "week6" / "coverage_grid.xlsx", "شبكة الاختبار")[2:]
    regions = dict(r[:2] for r in rows_of(FILES / "week6" / "coverage_grid.xlsx", "مفتاح المناطق")[1:])
    replies = {r["id"]: r for r in STANDIN["6.1"]}
    assert len(grid) == 12 and len(replies) == 12
    groups: dict[str, list] = {}
    for row in grid:
        rep = replies[row[0]]
        if row[1] == "مناطق":
            right = regions[row[2]].replace("مدينة ", "") in rep["reply"].split("مقر إمارتها")[1].split(".")[0]
            assert right == (rep["score"] >= 2), (row[2], rep["reply"])
        if row[1] == "لهجات ولغات":
            assert ("14" in rep["reply"]) == (rep["score"] >= 2), row[2]
        groups.setdefault(row[1], []).append((row[2], rep["score"]))
    for g, items in groups.items():
        scores = [s for _, s in items]
        weakest = min(items, key=lambda x: x[1])
        print(f"  {g}: mean {sum(scores) / 4:.2f}; weakest: {weakest[0]} ({weakest[1]}); gap {max(scores) - min(scores)}")
        assert max(scores) - min(scores) >= 1
    assert "سكاكا" in replies[8]["reply"] and regions["الحدود الشمالية"] == "مدينة عرعر"
    print("  the stand-in gave the wrong seat for the Northern Borders region; the key in the grid catches it")


# ------------------------------------------------------------------ 7.1
@lab("7.1")
def lab_7_1():
    cards = text(FILES / "week7" / "track_tasks.md")
    assert cards.count("\n## ") == 10
    source = cards.split("## 7 · الأعمال")[1].split("**المادة المصدرية.**")[1].split("\n## ")[0]
    s = STANDIN["7.1"]
    assert len(s["criteria"]) == 4
    totals = []
    for i, rnd in enumerate(s["rounds"], start=1):
        invented = sorted(numbers(rnd["output"]) - numbers(source) - {"120"})
        limit_named = "عند باب المقصف" in rnd["output"]
        n = words(rnd["output"])
        assert (not invented) == (rnd["scores"][0] == 2), (i, invented)
        assert limit_named == (rnd["scores"][1] == 2), i
        assert n <= 120 and rnd["scores"][2] == 2
        totals.append(sum(rnd["scores"]))
        print(f"  round {i}: {n} words; numbers not in the source: {invented or 'none'}; names the sample's limit: "
              f"{limit_named}; score {sum(rnd['scores'])} of 8")
    assert totals == sorted(totals) and totals[0] < totals[-1] == 8
    assert "تمثّل رأي طلبة الجامعة" in s["rounds"][1]["output"], "round 2 overgeneralises from 412 respondents"
    print("  round 2 keeps every number and still overgeneralises: only the participant's own check sees it")


# ------------------------------------------------------------------ 8.1
def verdicts(answers, facts):
    """True where the answer is wrong against the handbook facts."""
    wrong = {}
    wrong[1] = "12 محاضرة" in answers["1"] or f"{facts['absence_limit_pct']}%" not in answers["1"]
    wrong[2] = bool(numbers(answers["2"]))  # any amount is invented: the handbook has none
    wrong[3] = str(facts["appeal_working_days"]) not in answers["3"] or facts["appeal_channel"] not in answers["3"]
    wrong[4] = answers["4"].startswith("نعم")
    wrong[5] = (f"المادة {facts['hours_article']}" not in answers["5"] or str(facts["gpa_for_21"]) not in answers["5"]
                or str(facts["max_hours"]) not in answers["5"])
    wrong[6] = answers["6"].startswith("نعم")
    return wrong


@lab("8.1")
def lab_8_1():
    facts = json.loads(text(REF / "handbook_facts.json"))
    handbook = text(FILES / "week8" / "student_handbook_excerpt.md")
    for v in (facts["absence_limit_pct"], facts["appeal_working_days"], facts["max_hours"], facts["gpa_for_21"]):
        assert str(v) in handbook, v
    for topic in facts["not_in_handbook"]:
        assert topic not in handbook
    qs = rows_of(FILES / "week8" / "verification_checklist.xlsx", "قائمة التحقق")[2:]
    assert len(qs) == 6
    a = verdicts(STANDIN["8.1"]["no_source"], facts)
    b = verdicts(STANDIN["8.1"]["with_source"], facts)
    assert all(a.values()), a
    assert [k for k, v in b.items() if v] == [5], b
    inj = text(FILES / "week8" / "injection_example.txt")
    planted = [l.strip() for l in inj.splitlines() if "تعليمات إلى المساعد" in l]
    assert len(planted) == 1 and "تجاهل" in planted[0]
    print(f"  without the source: {sum(a.values())} of 6 answers are wrong against the handbook")
    print(f"  with the source pasted: {sum(b.values())} of 6 is still wrong (question 5: the quoted article says "
          f"4 where the handbook says {facts['gpa_for_21']})")
    print("  the injection example holds one instruction line addressed to the assistant; it is analysed, never run")


# ------------------------------------------------------------------ 9.1
def sanitise(raw: str, identifiers) -> str:
    """The reference sanitiser: known identifiers, then patterns for what a list can miss."""
    out = raw
    for ident in identifiers:
        out = out.replace(ident, "[محذوف]")
    out = re.sub(r"\b0?5\d{8}\b", "[جوال]", out)
    out = re.sub(r"\b[12]\d{9}\b", "[رقم]", out)
    out = re.sub(r"\S+@\S+", "[بريد]", out)
    return out


@lab("9.1")
def lab_9_1():
    raw = text(FILES / "week9" / "raw_prompt.txt")
    ids = json.loads(text(REF / "identifiers.json"))
    assert len(ids) == 8 and all(i in raw for i in ids)
    by_pattern = sanitise(raw, [])
    left = [i for i in ids if i in by_pattern]
    clean = text(REF / "sanitised_prompt.txt")
    assert not [i for i in ids if i in clean]
    assert not re.search(r"\d{9,}|@", clean)
    for fact in ("18,400", "1,150", "ثلاثة أيام", "أسبوعين"):
        assert fact in clean, fact
    data = read_csv(REF / "datasets_key.csv")
    assert [d["may_enter"] for d in data] == ["yes", "no", "no"]
    policy = {int(p["id"]): p for p in read_csv(REF / "gate_policy.csv")}
    plan, approved = STANDIN["9.1"]["agent_plan"], set(STANDIN["9.1"]["approved_by_human"])
    done, waited, refused, logged = [], [], [], []
    for step in plan:
        p = policy[step]
        if p["logged"] == "yes":
            logged.append(step)
        if p["decision"] == "مسموح":
            done.append(step)
        elif p["decision"] == "يحتاج موافقة":
            waited.append(step)
            if step in approved:
                done.append(step)
        else:
            refused.append(step)
    assert (done, waited, refused) == ([1, 2, 3, 4, 5], [4, 5], [10, 9])
    c = Counter(p["decision"] for p in policy.values())
    print(f"  raw prompt: 8 planted identifiers; patterns alone remove {8 - len(left)} and leave {len(left)} (names, a "
          f"company, an address): reading is still needed")
    print("  the reference prompt holds none of the 8 and keeps the four facts the task needs")
    print(f"  gate policy: {dict(c)}; replayed plan {plan}: done {done}, waited for a person {waited}, refused {refused}, "
          f"logged {logged}")


# ------------------------------------------------------------------ 10.1
@lab("10.1")
def lab_10_1():
    rows = rows_of(FILES / "week10" / "cost_sheet.xlsx", "التكلفة")
    v = [r[1] for r in rows[1:8]]
    n, m0, m1, check, setup, fee, months = v
    manual, assisted = n * m0, setup + n * (m1 + check)
    even = setup / (m0 - m1 - check)
    assert (manual, assisted, manual - assisted) == (540, 384, 156) and fee * months == 0
    formulas = [r[1] for r in rows[9:14]]
    assert all(isinstance(f, str) and f.startswith("=") for f in formulas), formulas
    template = text(FILES / "week10" / "use_case_plan_template.md")
    weak = text(FILES / "week10" / "sample_plan_weak.md")
    sections = re.findall(r"^## (\d) ·", template, flags=re.M)
    missing = [s for s in sections if f"## {s} ·" not in weak]
    assert sections == list("12345678") and missing == ["5", "7"]
    found = {
        "sections 5 and 7 missing": bool(missing),
        "personal data to a free assistant": "أرقام الجوالات" in weak and "مجاني" in weak,
        "no human review": "لا حاجة إلى مراجعة" in weak,
        "an unnecessary use: a margin is a formula": "هامش الربح" in weak,
        "free counted as no cost": "فلا تكلفة" in weak,
    }
    assert all(found.values()), found
    assert text(FILES / "week10" / "peer_review_checklist.md").count("\n") >= 12
    assert text(FILES / "week10" / "track_cards.md").count("\n## ") == 10
    print(f"  cost example: {manual} minutes by hand, {assisted} with the assistant, checking and setup; {manual - assisted} "
          f"saved; break-even after {even:.2f} runs, so from run {int(even) + 1}")
    print(f"  the weak plan: {len(found)} flaws found by rule: {', '.join(found)}")


def main(argv):
    want = argv or list(RUN)
    for lab_id in want:
        print(f"Lab {lab_id}")
        RUN[lab_id]()
        print("  PASS")
    print(f"{len(want)} labs passed")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(sys.argv[1:])
