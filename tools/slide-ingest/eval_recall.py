"""Slide Reference Pipeline v1 — Step 2 offline recall eval (see Idea/active/slide-reference-pipeline-plan.md).

  python eval_recall.py --subject RP --index "<LectureSlides>\\RP\\_index.json" --make-labels 30
      -> out/RP/labels.csv template (qid + question; fill source + page by hand, page may be "38|39")
  python eval_recall.py --subject RP --index "<LectureSlides>\\RP\\_index.json"
      -> top-1 / top-5 recall (page-level and lecture-level) on labeled rows

Questions: out/<SUBJ>/questions.json (GAS getQuestions export).
The BM25 SPEC block below must stay identical to DATABASE js/slide-ref.js (Step 5).
"""
import argparse
import ast
import csv
import json
import math
import random
import re
from collections import Counter
from pathlib import Path

from ingest_slides import norm_title

HERE = Path(__file__).resolve().parent
TOPIC_OVERRIDES = {}  # set in main() from _index.json meta.topic_map (same map the browser gets)

# ---- SPEC (mirror in js/slide-ref.js) ----
K1, B = 1.2, 0.75
BIGRAM_WEIGHT = 1.0
ANSWER_REPEAT = 2          # query = stem + 2x correct answer, no distractors
CAT_BOOST = 1.5            # question category topic ~ page source lecture title
CAT_MATCH_MIN = 0.6        # token-set Jaccard (topic vs lecture title) for the boost; + _index.json meta.topic_map
TOP_K = 5
STOPWORDS = set("""
a an the and or of in on at to for from by with without as is are was were be been being this that these those
it its which what who whom whose when where why how than then there their they them he she his her we our you your
not no nor but if so such can could may might will would should shall do does did has have had having also
most least more less very all any each both few other some same only own into over under about after before
between during through above below up down out off again further once here following true false except
patient patients year years old case cause causes caused likely diagnosis statement
correct incorrect best choice answer question associated regarding
""".split())
# ------------------------------------------


def tokenize(text):
    """EN unigram + bigram; Thai/punctuation/digit-only dropped."""
    words = [w for w in re.findall(r"[a-z][a-z0-9]*", text.lower()) if len(w) > 1 and w not in STOPWORDS]
    return words + [f"{a}_{b}" for a, b in zip(words, words[1:])]


def build_bm25(pages):
    docs = [Counter(tokenize(p["title"] + "\n" + p["slide_text"])) for p in pages]
    lens = [sum(d.values()) for d in docs]
    df = Counter(t for d in docs for t in d)
    return {"docs": docs, "lens": lens, "avg": sum(lens) / len(lens), "df": df, "n": len(docs)}


def score(query_tokens, ix, boosts):
    q = Counter(query_tokens)
    out = []
    for i, d in enumerate(ix["docs"]):
        s = 0.0
        for t, qf in q.items():
            f = d.get(t)
            if not f:
                continue
            idf = math.log(1 + (ix["n"] - ix["df"][t] + 0.5) / (ix["df"][t] + 0.5))
            w = BIGRAM_WEIGHT if "_" in t else 1.0
            s += w * qf * idf * f * (K1 + 1) / (f + K1 * (1 - B + B * ix["lens"][i] / ix["avg"]))
        out.append(s * boosts[i])
    return out


def question_topics(q):
    """"['RP_51MCQ1', 'RP_ANA_Anatomy of pelvis']" -> ['Anatomy of pelvis'] (exam-set ids dropped)."""
    try:
        cats = ast.literal_eval(q["category"]) if isinstance(q["category"], str) else q["category"]
    except (ValueError, SyntaxError):
        cats = [q["category"]]
    return [re.sub(r"^.*_(?:ANA|PHYSIO|MICRO|PATHO|PHARM|CLINICAL|BIOCHEM|PARASITO|RADIO|LAB)_", "", c)
            for c in cats if re.search(r"_(?:ANA|PHYSIO|MICRO|PATHO|PHARM|CLINICAL|BIOCHEM|PARASITO|RADIO|LAB)_", c)]


def topic_tokens(s):
    s = norm_title(s).replace("reproductive system", "rp")
    return set(re.sub(r"\bsexual(ly)?\b", "sex", s).split()) - {"of", "and"}


def topic_sources(q, sources, overrides=None):
    """Lecture sources a question's category topic points at: override list, else best Jaccard >= CAT_MATCH_MIN."""
    overrides = TOPIC_OVERRIDES if overrides is None else overrides
    hit = set()
    for t in question_topics(q):
        if overrides and t.lower() in overrides:
            hit.update(s for s in overrides[t.lower()] if s in sources)
            continue
        tt = topic_tokens(t)
        sims = {s: len(tt & topic_tokens(s)) / len(tt | topic_tokens(s)) for s in sources}
        best = max(sims.values(), default=0)
        if best >= CAT_MATCH_MIN:
            hit.update(s for s, v in sims.items() if v == best)
    return hit


def load_topic_overrides(meta):
    # meta.topic_map = {"<category topic>": ["<source>", ...]} — topic keys matched case-insensitively
    return {k.lower(): v for k, v in meta.get("topic_map", {}).items()}


def query_tokens(q):
    """stem + ANSWER_REPEAT x answer, tokenized per part so no bigram spans the join."""
    stem = re.sub(r"^\s*\d+\s*[.)]\s*", "", q["problem"])
    return tokenize(stem) + tokenize(q["answer"]) * ANSWER_REPEAT


def rank(q, pages, ix, boost=True):
    srcs = topic_sources(q, {p["source"] for p in pages}) if boost else set()
    boosts = [CAT_BOOST if p["source"] in srcs else 1.0 for p in pages]
    s = score(query_tokens(q), ix, boosts)
    return sorted(range(len(pages)), key=lambda i: s[i], reverse=True)[:TOP_K], s


def make_labels(qs, pages, n, dest):
    sources = {p["source"] for p in pages}
    pool = [q for q in qs if topic_sources(q, sources) and q.get("problem", "").strip()]
    random.seed(52)
    by_src = {}
    for q in pool:
        by_src.setdefault(sorted(topic_sources(q, sources))[0], []).append(q)
    picked, srcs = [], sorted(by_src)
    while len(picked) < n and any(by_src.values()):  # round-robin across lectures
        for s in srcs:
            if by_src[s] and len(picked) < n:
                picked.append(by_src[s].pop(random.randrange(len(by_src[s]))))
    with dest.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["qid", "topic_lecture_hint", "problem", "answer", "source", "page"])
        for q in picked:
            w.writerow([q["questionId"], "; ".join(sorted(topic_sources(q, sources))),
                        q["problem"][:400], q["answer"][:200], "", ""])
    print(f"mappable questions {len(pool)}/{len(qs)} | wrote {len(picked)} rows -> {dest}")


def evaluate(qs, pages, ix, labels_path):
    qmap = {q["questionId"]: q for q in qs}
    rows = []
    with labels_path.open(encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            if not row["source"].strip() or not row["page"].strip():
                continue
            if row["qid"] not in qmap:
                print("  qid not in export:", row["qid"])
                continue
            rows.append((qmap[row["qid"]], row["source"].strip(), {int(x) for x in row["page"].split("|")}))
    if not rows:
        print("no labeled rows (fill source + page in labels.csv)")
        return
    # boost=off approximates the ~1/3 of real questions whose category maps to no lecture
    for boost in (True, False):
        n = len(rows)
        p1 = p5 = l1 = l5 = 0
        misses = []
        for q, src, want in rows:
            top, _ = rank(q, pages, ix, boost)
            hits_p = [pages[i]["source"] == src and pages[i]["page_no"] in want for i in top]
            hits_l = [pages[i]["source"] == src for i in top]
            p1 += hits_p[0]; p5 += any(hits_p); l1 += hits_l[0]; l5 += any(hits_l)
            if not any(hits_p):
                misses.append((q["questionId"], src, sorted(want),
                               [f"{pages[i]['lecture']}p{pages[i]['page_no']}" for i in top]))
        print(f"boost {'ON ' if boost else 'OFF'} | labeled {n} | page top-1 {p1/n:.0%} top-5 {p5/n:.0%} | "
              f"lecture top-1 {l1/n:.0%} top-5 {l5/n:.0%} | v2 gate (page top-5 >= 80%): "
              f"{'PASS' if p5/n >= 0.8 else 'FAIL'}")
        if boost:
            for m in misses:
                print("  miss", *m)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--subject", required=True)
    ap.add_argument("--index", required=True, type=Path)
    ap.add_argument("--make-labels", type=int)
    a = ap.parse_args()
    idx = json.loads(a.index.read_text(encoding="utf-8"))
    TOPIC_OVERRIDES.update(load_topic_overrides(idx["meta"]))
    out = HERE / "out" / a.subject
    pages = idx["pages"]
    qs = json.loads((out / "questions.json").read_text(encoding="utf-8"))
    if a.make_labels:
        make_labels(qs, pages, a.make_labels, out / "labels.csv")
        return
    evaluate(qs, pages, build_bm25(pages), out / "labels.csv")


if __name__ == "__main__":
    main()
