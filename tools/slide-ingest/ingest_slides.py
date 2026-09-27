"""Slide Reference Pipeline v1 — Step 1 ingest (see Idea/active/slide-reference-pipeline-plan.md).

Usage:
  python ingest_slides.py --subject RP --src "D:\\00 POND\\University\\Y3\\RP" --pairs-only

  python ingest_slides.py --subject RP --src "..." --out "<Drive-synced>\\LectureSlides\\RP"

--pairs-only: pair pdf (src/lecture) <-> md (src/markdown), write out/<SUBJ>/pairs.report.txt, stop.
default: render paired pages -> <out>/<SUBJ>__<tag>__pNNN.webp (skip existing) + <out>/_index.json
({"meta": {"subject", "topic_map"}, "pages": [row, ...]}; topic_map from topic_map.override.json).
Range sections fan out: each page gets its own row sharing the section text.
"""
import argparse
import difflib
import json
import re
import sys
import time
from pathlib import Path

import fitz  # PyMuPDF
from PIL import Image

HERE = Path(__file__).resolve().parent
CODE_RE = re.compile(r"^\s*(L\d+(?:-\d+)?)\b", re.I)
# '## [n. | n–m. ]<title> — หน้า <spec>'  spec = '3' | '1–2' | '3, 8–9'
PAGE_RE = re.compile(r"^##\s+(?:\d+(?:\s*[–-]\s*\d+)?\.\s+)?(.+?)\s+—\s+หน้า\s+([\d\s,–-]+?)\s*$")
# inner per-page marker inside a range section: '### หน้า 6 — …', '#### หน้า 27: …', '**หน้า 14 (…):**'
MARK_RE = re.compile(r"^(?:#{3,4}\s+|\*\*)หน้า\s+(\d+(?:\s*[–-]\s*\d+)?)(?!\d)")
SPEC_PART_RE =re.compile(r"^(\d+)(?:\s*[–-]\s*(\d+))?$")
H2_RE = re.compile(r"^##\s")  # exactly '## ', not '###'
MIN_SIM = 0.6
RENDER_WIDTH = 1200
WEBP_Q = 70
NOTES_CAP = 20000


def lecture_code(stem):
    m = CODE_RE.match(stem)
    return m.group(1).upper() if m else None


def norm_title(stem):
    s = CODE_RE.sub("", stem).lower()
    s = re.sub(r"_compressed\b", "", s)
    s = re.sub(r"[^a-z0-9\u0e00-\u0e7f]+", " ", s)
    return " ".join(s.split())


def similarity(a, b):
    return difflib.SequenceMatcher(None, norm_title(a), norm_title(b)).ratio()


def expand_spec(spec):
    """'6–8, 11' -> [6, 7, 8, 11]; None if malformed."""
    pages = []
    for part in spec.split(","):
        m = SPEC_PART_RE.match(part.strip())
        if not m:
            return None
        a, b = int(m.group(1)), int(m.group(2) or m.group(1))
        if b < a:
            return None
        pages.extend(range(a, b + 1))
    return pages


def md_stats(path):
    parsed, multi, skipped, unparsed, pages = 0, 0, 0, [], set()
    for line in path.read_text(encoding="utf-8").splitlines():
        if not H2_RE.match(line):
            continue
        if "Topic Checklist" in line or "หน้า" not in line:
            skipped += 1
            continue
        m = PAGE_RE.match(line)
        spec_pages = expand_spec(m.group(2)) if m else None
        if not spec_pages:
            unparsed.append(line.strip())
            continue
        parsed += 1
        if len(spec_pages) > 1:
            multi += 1
        pages.update(spec_pages)
    return {"max_page": max(pages, default=0), "parsed": parsed, "multi": multi,
            "skipped": skipped, "unparsed": unparsed, "pages": pages}


def parse_sections(path):
    """-> [{title, pages, lines:[(bucket, tag)]}]; one per page heading.
    tag = set of pages from an inner marker (### หน้า N / #### หน้า N / **หน้า N), None = shared."""
    sections, cur, sub, tag = [], None, None, None

    for line in path.read_text(encoding="utf-8").splitlines():
        if line.startswith("# ") or H2_RE.match(line):
            cur, sub, tag = None, None, None  # any #/## ends the current page section
            if H2_RE.match(line) and "Topic Checklist" not in line and "หน้า" in line:
                m = PAGE_RE.match(line)
                pages = expand_spec(m.group(2)) if m else None
                if pages:
                    cur = {"title": m.group(1).strip(), "pages": pages, "lines": []}
                    sections.append(cur)
            continue
        if cur is None:
            continue
        mk = MARK_RE.match(line)
        if line.startswith("### "):
            # '### หน้า N — …' = per-page slide block (L11 style); other ### = new subsection, shared
            sub = "slide" if (mk or "เนื้อหาจากสไลด์" in line) else "notes"
            tag = set(expand_spec(mk.group(1)) or []) if mk else None
            if sub == "slide" and not mk:
                continue  # plain '### เนื้อหาจากสไลด์' label is not content
        elif mk:
            tag = set(expand_spec(mk.group(1)) or [])
        cur["lines"].append(("slide" if sub == "slide" else "notes", frozenset(tag) if tag else None, line))
    return sections


def page_text(sec, p):
    """(slide_text, notes_md) for page p: shared lines + lines tagged with p."""
    out = {"slide": [], "notes": []}
    for bucket, tag, line in sec["lines"]:
        if tag is None or p in tag:
            out[bucket].append(line)
    return "\n".join(out["slide"]).strip(), "\n".join(out["notes"]).strip()[:NOTES_CAP]


def render_page(page, dest):
    zoom = RENDER_WIDTH / page.rect.width
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    Image.frombytes("RGB", (pix.width, pix.height), pix.samples).save(dest, "WEBP", quality=WEBP_Q)


def lecture_tags(pairs):
    """Filename tag per pair: lecture code, + title slug when a code is shared (two L17-18)."""
    codes = [lecture_code(md.stem) for _, md, _, _ in pairs]
    tags = {}
    for (pdf, md, _, _), code in zip(pairs, codes):
        slug = "-".join(norm_title(md.stem).split()[:2])
        tags[md.name] = code if code and codes.count(code) == 1 else f"{code or 'NA'}-{slug}"
    return tags


def render_all(subject, src, out):
    t0 = time.time()
    pdfs = sorted((src / "lecture").glob("*.pdf"))
    mds = sorted((src / "markdown").glob("*.md"))
    pairs, _, _, _ = pair(pdfs, mds, load_overrides(HERE / "pairs.override.json", subject))
    pairs.sort(key=lambda p: p[1].name)
    tags = lecture_tags(pairs)
    out.mkdir(parents=True, exist_ok=True)

    index, made, skipped, warns = [], 0, 0, []
    text_srcs = {"md": 0, "md+pdf": 0, "pdf": 0, "title": 0}
    for pdf, md, _, _ in pairs:
        tag = tags[md.name]
        source = re.sub(r"_compressed$", "", md.stem)
        with fitz.open(pdf) as doc:
            for sec in parse_sections(md):
                for p in sec["pages"]:
                    if p > doc.page_count:
                        warns.append(f"{md.name}: หน้า {p} > pdf {doc.page_count} pages, skipped")
                        continue
                    img = f"{subject}__{tag}__p{p:03d}.webp"
                    dest = out / img
                    if dest.exists():
                        skipped += 1
                    else:
                        render_page(doc[p - 1], dest)
                        made += 1
                    slide_text, notes_md = page_text(sec, p)
                    text_src = "md"
                    pdf_text = doc[p - 1].get_text().strip()
                    if not slide_text:  # safety net: pdf text layer, then section title
                        slide_text, text_src = pdf_text, "pdf"
                        if not slide_text:
                            slide_text, text_src = sec["title"], "title"
                    elif len(sec["pages"]) > 1 and pdf_text:
                        # multi-page block: md text may be shared across pages -> add page-specific words
                        slide_text, text_src = slide_text + "\n" + pdf_text, "md+pdf"
                    text_srcs[text_src] += 1
                    index.append({"source": source, "lecture": lecture_code(md.stem), "page_no": p,
                                  "title": sec["title"], "slide_text": slide_text, "slide_text_src": text_src,
                                  "notes_md": notes_md, "image_file": img, "source_type": "slide"})
        print(f"  {tag}: done ({time.time() - t0:.0f}s)")

    # meta.topic_map ships the local category-topic -> lecture override to the browser (Step 5 boost)
    meta = {"subject": subject, "topic_map": load_overrides(HERE / "topic_map.override.json", subject)}
    (out / "_index.json").write_text(json.dumps({"meta": meta, "pages": index}, ensure_ascii=False, indent=1),
                                     encoding="utf-8")
    mb = sum(f.stat().st_size for f in out.glob("*.webp")) / 1e6
    print(f"rows {len(index)} | rendered {made} | existing {skipped} | webp total {mb:.1f}MB | "
          f"slide_text src {text_srcs} | warn {len(warns)} | {time.time() - t0:.0f}s")
    for w in warns:
        print("  WARN", w)


def pdf_pages(path):
    with fitz.open(path) as doc:
        return doc.page_count


def load_overrides(path, subject):
    # {"RP": {"<md filename>": "<pdf filename>" | null}}  null = force unpaired
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8")).get(subject, {})


def pair(pdfs, mds, overrides):
    pairs, notes = [], []
    pdf_left, md_left = {p.name: p for p in pdfs}, {m.name: m for m in mds}

    for md_name, pdf_name in overrides.items():
        md = md_left.pop(md_name, None)
        if md is None:
            notes.append(f"override md not found: {md_name}")
            continue
        if pdf_name is None:
            notes.append(f"override forced unpaired: {md_name}")
            continue
        pdf = pdf_left.pop(pdf_name, None)
        if pdf is None:
            notes.append(f"override pdf not found: {pdf_name} (md {md_name} left unpaired)")
            continue
        pairs.append((pdf, md, similarity(pdf.stem, md.stem), "override"))

    # candidates: same lecture code (both None allowed), sim >= MIN_SIM
    cands = []
    for md in md_left.values():
        for pdf in pdf_left.values():
            if lecture_code(md.stem) != lecture_code(pdf.stem):
                continue
            sim = similarity(pdf.stem, md.stem)
            if sim >= MIN_SIM:
                exact = pdf.stem.lower() == md.stem.lower()
                cands.append((sim, exact, pdf, md))
    # greedy one-to-one; raw-stem exact match breaks ties (e.g. *_compressed twins)
    cands.sort(key=lambda c: (c[0], c[1]), reverse=True)
    for sim, exact, pdf, md in cands:
        if pdf.name not in pdf_left or md.name not in md_left:
            continue
        rivals = [c for c in cands if c[3] is md and c[2] is not pdf and abs(c[0] - sim) < 1e-9]
        how = "auto"
        if rivals:
            how = "auto-tiebreak(exact stem)" if exact else "auto-TIE(unresolved, first wins)"
            notes.append(f"tie for {md.name}: chose {pdf.name} over {', '.join(r[2].name for r in rivals)}")
        del pdf_left[pdf.name], md_left[md.name]
        pairs.append((pdf, md, sim, how))

    return pairs, list(pdf_left.values()), list(md_left.values()), notes


def pairs_report(subject, src):
    pdfs = sorted((src / "lecture").glob("*.pdf"))
    mds = sorted((src / "markdown").glob("*.md"))
    overrides = load_overrides(HERE / "pairs.override.json", subject)
    pairs, un_pdf, un_md, notes = pair(pdfs, mds, overrides)
    pairs.sort(key=lambda p: p[1].name)

    out = [f"# pairs.report — {subject}", f"src: {src}", f"pdf: {len(pdfs)}  md: {len(mds)}  paired: {len(pairs)}", ""]
    warns = 0
    for pdf, md, sim, how in pairs:
        n_pdf = pdf_pages(pdf)
        st = md_stats(md)
        flag = ""
        if st["max_page"] > n_pdf:
            flag = f"  WARN md max page {st['max_page']} > pdf pages {n_pdf}"
            warns += 1
        out += [
            f"[{lecture_code(md.stem) or '-'}] sim={sim:.2f} ({how}){flag}",
            f"  pdf: {pdf.name}  ({n_pdf} pages)",
            f"  md : {md.name}  (max หน้า {st['max_page']}, headings {st['parsed']}, multi-page {st['multi']}, "
            f"skipped ## {st['skipped']}, unparsed ## {len(st['unparsed'])})",
            f"  pages covered {len(st['pages'])}/{n_pdf}",
        ]
        missing = sorted(set(range(1, n_pdf + 1)) - st["pages"])
        if missing:
            out.append(f"      no md text: {missing}")
        out += [f"      unparsed: {u}" for u in st["unparsed"][:5]]
        if len(st["unparsed"]) > 5:
            out.append(f"      ... +{len(st['unparsed']) - 5} more")
    out += ["", f"UNPAIRED pdf ({len(un_pdf)}) — skipped:"] + [f"  {p.name}" for p in un_pdf]
    out += ["", f"UNPAIRED md ({len(un_md)}) — skipped:"] + [f"  {m.name}" for m in un_md]
    if notes:
        out += ["", "NOTES:"] + [f"  {n}" for n in notes]

    dest = HERE / "out" / subject / "pairs.report.txt"
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_text("\n".join(out) + "\n", encoding="utf-8")
    print(f"paired {len(pairs)} | unpaired pdf {len(un_pdf)} md {len(un_md)} | page WARN {warns} | notes {len(notes)}")
    print(f"report: {dest}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--subject", required=True)
    ap.add_argument("--src", required=True, type=Path)
    ap.add_argument("--out", type=Path, help="Drive-synced LectureSlides/<SUBJ> dir (render mode)")
    ap.add_argument("--pairs-only", action="store_true")
    a = ap.parse_args()
    if a.pairs_only:
        pairs_report(a.subject, a.src)
        return
    if not a.out:
        sys.exit("--out required for render mode")
    render_all(a.subject, a.src, a.out)


if __name__ == "__main__":
    main()
