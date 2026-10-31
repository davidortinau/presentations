#!/usr/bin/env python3
"""Prepare a slide deck for public sharing.

  deck.py inspect DECK.pptx --out REVIEW.md [--terms TERMS.txt]
      Write a review file: properties, sensitivity labels, comments, hidden slides, embedded media,
      and every slide's text and speaker notes, with flagged terms listed first. The review file
      holds the deck's private notes, so keep it out of this repository.

  deck.py apply DECK.pptx EDITS.json OUT.pptx
      Write a public copy. EDITS.json can hold:
        "notes":       [{"slide": 3, "old": "exact text", "new": "replacement"}]
        "clear_notes": [5, 7]                 remove all notes on these slides
        "slide_text":  [{"slide": 2, "old": "exact text", "new": "replacement"}]   text inside one run
        "paragraph_text": [{"slide": 2, "old": "...", "new": "..."}]   text split across runs
        "drop_slides": [12]                   remove these slides (usually hidden ones)
        "core":        {"title": "...", "subject": "...", "creator": "David Ortinau", "keywords": ""}
      Every "old" text must match exactly once, or nothing is written. Comments, comment authors,
      and the company and manager properties are always removed. Parts that no edit touches are
      copied byte for byte.

  deck.py pdf DECK.pptx OUT.pdf [--engine auto|powerpoint|libreoffice] [--max-mb 20]
      Export the slides (no notes) to PDF. PowerPoint for Mac gives the best fidelity; LibreOffice
      is the fallback. A PDF over --max-mb gets its images resampled to 150 dpi.

  deck.py thumb DECK.pdf OUT.jpg [--width 1280]
      Render the first slide as a JPEG.
"""
import argparse
import fnmatch
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from xml.etree import ElementTree as ET

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    "cu": "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}

# Terms that need a human look before a deck goes public. Customer and partner names go in a
# private --terms file, never in this list.
FLAGS = [
    ("confidentiality", r"\b(confidential|internal[- ]only|internal use|under nda|\bnda\b|do not (share|distribute|forward)|don'?t share|not public|unannounced|embargo(ed)?|privileged)\b"),
    ("internal", r"\binternal\b"),
    ("roadmap or dates", r"\b(roadmap|FY\s?\d{2}|CY\s?\d{2}|[QH][1-4]\s?(FY|CY)?\s?'?\d{2,4}|ship date|planned for|eta\b)"),
    ("internal host", r"(sharepoint\.com|microsoft-my\.sharepoint|teams\.microsoft\.com|dev\.azure\.com|\.visualstudio\.com|eng\.ms|\.corp\.|corpnet|msit|intranet|1drv\.ms|onedrive\.live|stream\.microsoft|microsoftstream|loop\.microsoft|aka\.ms/[A-Za-z0-9._-]+|\bIcM\b|\bADO\b)"),
    ("email", r"[\w.+-]+@[\w-]+\.[\w.]+"),
    ("local path", r"(/Users/|~/|\.\./\.\./|[A-Z]:\\|\bdavidos\b|\bProjects/|\breports/|\bcontext/)"),
    ("metric", r"\b\d[\d,.]*\s?[kKmMbB%]?\s+(users|developers|devs|customers|apps|projects|downloads|installs|MAU|DAU|WAU|seats)\b"),
    ("people and process", r"\b(director review|exec(utive)? review|leadership|headcount|reorg|budget|okr|kpi|telemetry|kusto|survey results)\b"),
    ("placeholder", r"\b(TODO|TBD|FIXME|XXX|lorem ipsum)\b"),
]
PUBLIC_LABELS = {"public", "general", "non-business", "personal"}
COMMENT_REL_TYPES = ("/comments", "/commentAuthors", "/authors")


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def text_of(xml_bytes):
    """All a:t text in a part, one line per paragraph."""
    root = ET.fromstring(xml_bytes)
    paras = []
    for p in root.iter(f"{{{NS['a']}}}p"):
        t = "".join(x.text or "" for x in p.iter(f"{{{NS['a']}}}t"))
        if t.strip():
            paras.append(t)
    return "\n".join(paras)


def rels_of(z, part):
    d, b = os.path.split(part)
    rp = f"{d}/_rels/{b}.rels"
    if rp not in z.namelist():
        return []
    root = ET.fromstring(z.read(rp))
    out = []
    for r in root.findall("rel:Relationship", NS):
        tgt = r.get("Target")
        if r.get("TargetMode") == "External":
            full = tgt
        else:
            full = os.path.normpath(os.path.join(d, tgt)).replace("\\", "/")
        out.append((r.get("Id"), r.get("Type"), full, r.get("TargetMode") == "External"))
    return out


def slide_parts(z):
    """Slide part names in presentation order."""
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rid_to_part = {rid: part for rid, _, part, _ in rels_of(z, "ppt/presentation.xml")}
    ids = pres.find("p:sldIdLst", NS)
    if ids is None:
        return []
    return [rid_to_part[s.get(f"{{{NS['r']}}}id")] for s in ids.findall("p:sldId", NS)]


def props(z):
    out = {}
    if "docProps/core.xml" in z.namelist():
        root = ET.fromstring(z.read("docProps/core.xml"))
        for el in root:
            out["core." + el.tag.split("}")[1]] = (el.text or "").strip()
    if "docProps/app.xml" in z.namelist():
        root = ET.fromstring(z.read("docProps/app.xml"))
        for tag in ("Company", "Manager", "Template", "Application", "Slides", "Notes", "HiddenSlides", "MMClips"):
            el = root.find(f"ep:{tag}", NS)
            if el is not None:
                out["app." + tag] = (el.text or "").strip()
    custom = {}
    if "docProps/custom.xml" in z.namelist():
        root = ET.fromstring(z.read("docProps/custom.xml"))
        for p in root.findall("cu:property", NS):
            val = "".join(x.text or "" for x in p)
            custom[p.get("name")] = val
    return out, custom


def labels_of(custom):
    """Sensitivity label names from MSIP_Label_<guid>_Name properties."""
    return sorted({v for k, v in custom.items() if k.startswith("MSIP_Label_") and k.endswith("_Name")})


def comments_of(z):
    out = []
    for n in z.namelist():
        if n.startswith("ppt/comments/") and n.endswith(".xml"):
            out.append((n, text_of(z.read(n))))
    return out


def load_terms(path):
    if not path:
        return []
    terms = []
    for line in open(path, encoding="utf-8"):
        line = line.strip()
        if line and not line.startswith("#"):
            terms.append(line)
    return terms


def find_flags(text, terms):
    hits = []
    for name, pat in FLAGS:
        for m in re.finditer(pat, text, re.I):
            hits.append((name, m.group(0), text[max(0, m.start() - 60): m.end() + 60].replace("\n", " ")))
    for t in terms:
        for m in re.finditer(r"\b" + re.escape(t) + r"\b", text, re.I):
            hits.append(("named term", m.group(0), text[max(0, m.start() - 60): m.end() + 60].replace("\n", " ")))
    return hits


def cmd_inspect(a):
    terms = load_terms(a.terms)
    z = zipfile.ZipFile(a.deck)
    names = z.namelist()
    core, custom = props(z)
    labels = labels_of(custom)
    slides = slide_parts(z)
    rows, flags, hidden = [], [], []
    for i, part in enumerate(slides, 1):
        xml = z.read(part)
        is_hidden = ET.fromstring(xml).get("show") == "0"
        if is_hidden:
            hidden.append(i)
        stext = text_of(xml)
        notes = ""
        media = []
        for _, rtype, target, external in rels_of(z, part):
            if rtype.endswith("/notesSlide") and target in names:
                notes = text_of(z.read(target))
            elif external:
                media.append(("link", target, 0))
            elif target in names and target.startswith(("ppt/media/", "ppt/embeddings/")):
                media.append((rtype.rsplit("/", 1)[-1], target, z.getinfo(target).file_size))
        for where, txt in (("slide", stext), ("notes", notes)):
            for f in find_flags(txt, terms):
                flags.append((i, where) + f)
        rows.append((i, is_hidden, stext, notes, media))
    comments = comments_of(z)
    for n, txt in comments:
        for f in find_flags(txt, terms):
            flags.append(("-", "comment " + os.path.basename(n)) + f)
    media_all = [(n, z.getinfo(n).file_size) for n in names if n.startswith("ppt/media/") and not n.endswith("/")]
    embeds = [(n, z.getinfo(n).file_size) for n in names if n.startswith("ppt/embeddings/") and not n.endswith("/")]
    videos = [m for m in media_all if re.search(r"\.(mp4|mov|m4v|wmv|avi|mkv)$", m[0], re.I)]
    risky_labels = [l for l in labels if l.strip().lower() not in PUBLIC_LABELS]

    mb = lambda b: f"{b / 1e6:.1f} MB"
    L = []
    L.append(f"# Review: {os.path.basename(a.deck)}\n")
    L.append("| Item | Value |\n|---|---|")
    L.append(f"| File | `{a.deck}` |")
    L.append(f"| SHA-256 | `{sha256(a.deck)}` |")
    L.append(f"| Size | {mb(os.path.getsize(a.deck))} |")
    L.append(f"| Slides | {len(slides)} (hidden: {', '.join(map(str, hidden)) or 'none'}) |")
    L.append(f"| Slides with notes | {sum(1 for r in rows if r[3].strip())} |")
    for k in ("core.title", "core.subject", "core.creator", "core.lastModifiedBy", "core.keywords", "core.description", "core.created", "core.modified", "app.Company", "app.Manager", "app.Template"):
        if core.get(k):
            L.append(f"| {k} | {core[k][:120]} |")
    L.append(f"| Sensitivity labels | {', '.join(labels) or 'none'}{'  **ABOVE GENERAL: do not publish**' if risky_labels else ''} |")
    L.append(f"| Comments | {len(comments)} part(s) |")
    L.append(f"| Media | {len(media_all)} files, {mb(sum(s for _, s in media_all))}; videos: {len(videos)} ({mb(sum(s for _, s in videos))}) |")
    L.append(f"| Embedded objects | {len(embeds)}{': ' + ', '.join(os.path.basename(n) for n, _ in embeds) if embeds else ''} |")
    L.append("")
    L.append(f"## Flags ({len(flags)})\n")
    if flags:
        L.append("| Slide | Where | Kind | Match | Context |\n|---|---|---|---|---|")
        for s, where, kind, match, ctx in flags:
            ctx = ctx.replace("|", "\\|")
            L.append(f"| {s} | {where} | {kind} | `{match}` | {ctx} |")
    else:
        L.append("None.")
    if comments:
        L.append("\n## Comments\n")
        for n, txt in comments:
            L.append(f"- `{n}`: {txt[:300]}")
    L.append("\n## Slides\n")
    for i, is_hidden, stext, notes, media in rows:
        L.append(f"### Slide {i}{' (hidden)' if is_hidden else ''}\n")
        L.append("Text:\n```\n" + (stext or "(none)") + "\n```")
        L.append("Notes:\n```\n" + (notes or "(none)") + "\n```")
        if media:
            L.append("Media: " + "; ".join(f"{k} {os.path.basename(t)}{' ' + mb(s) if s else ''}" for k, t, s in media))
        L.append("")
    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    open(a.out, "w", encoding="utf-8").write("\n".join(L))
    summary = dict(slides=len(slides), hidden=hidden, flags=len(flags), comments=len(comments),
                   labels=labels, risky_labels=risky_labels, videos=len(videos), embeds=len(embeds))
    print(json.dumps(summary))


def _strip_comments(parts):
    """Remove comment parts, comment authors, and the relationships and content types that point at them."""
    gone = {n for n in parts if n.startswith("ppt/comments/") or n in ("ppt/commentAuthors.xml", "ppt/authors.xml")}
    for n in list(parts):
        if n.endswith(".rels") and n not in gone:
            root = ET.fromstring(parts[n])
            drop = [r for r in root.findall("rel:Relationship", NS) if r.get("Type", "").endswith(COMMENT_REL_TYPES)]
            if drop:
                for r in drop:
                    root.remove(r)
                parts[n] = _xml(root, NS["rel"])
    for n in list(parts):
        if n.startswith("ppt/comments/_rels/"):
            gone.add(n)
    for n in gone:
        parts.pop(n, None)
    if gone and "[Content_Types].xml" in parts:
        root = ET.fromstring(parts["[Content_Types].xml"])
        for o in root.findall("ct:Override", NS):
            if o.get("PartName", "").lstrip("/") in gone:
                root.remove(o)
        parts["[Content_Types].xml"] = _xml(root, NS["ct"])
    return sorted(n for n in gone if not n.endswith(".rels"))


def _xml(root, default_ns):
    ET.register_namespace("", default_ns)
    for k, v in NS.items():
        if v != default_ns and k not in ("rel", "ct", "cu", "ep"):
            ET.register_namespace(k, v)
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + ET.tostring(root)


def _set_core(data, core):
    for key, val in core.items():
        ns = "cp" if key in ("keywords", "lastModifiedBy", "category") else "dc"
        tag = f"{ns}:{key}"
        val_x = (val or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").encode()
        pat = re.compile(rb"<" + tag.encode() + rb"(\s[^>]*)?(/>|>.*?</" + tag.encode() + rb">)", re.S)
        if pat.search(data):
            data = pat.sub(b"<" + tag.encode() + b">" + val_x + b"</" + tag.encode() + b">", data, count=1)
        elif val:
            data = data.replace(b"</cp:coreProperties>", b"<" + tag.encode() + b">" + val_x + b"</" + tag.encode() + b"></cp:coreProperties>")
    return data


def _clear_app(data):
    for tag in (b"Company", b"Manager"):
        data = re.sub(rb"<" + tag + rb">.*?</" + tag + rb">", b"<" + tag + b"></" + tag + b">", data, flags=re.S)
    return data


def cmd_apply(a):
    from pptx import Presentation

    edits = json.load(open(a.edits, encoding="utf-8"))
    prs = Presentation(a.deck)
    slides = list(prs.slides)
    changed = {}
    for e in edits.get("notes", []):
        n = e["slide"]
        cur = changed.get(n, slides[n - 1].notes_slide.notes_text_frame.text if slides[n - 1].has_notes_slide else "")
        count = cur.count(e["old"])
        if count != 1:
            sys.exit(f"slide {n} notes: expected 1 match, found {count}: {e['old'][:80]!r}")
        changed[n] = cur.replace(e["old"], e["new"])
    for n in edits.get("clear_notes", []):
        changed[n] = ""
    for n, text in changed.items():
        slides[n - 1].notes_slide.notes_text_frame.text = text

    # Slide text that PowerPoint split across several runs (common for emails and links): match the
    # paragraph's joined text, then keep the first run's formatting for the whole replaced paragraph.
    para_touched = set()
    for e in edits.get("paragraph_text", []):
        slide = slides[e["slide"] - 1]
        hits = 0

        def frames(shapes):
            for sh in shapes:
                if sh.shape_type == 6:  # group
                    yield from frames(sh.shapes)
                elif sh.has_text_frame:
                    yield sh.text_frame

        for tf in frames(slide.shapes):
            for p in tf.paragraphs:
                runs = list(p.runs)
                full = "".join(r.text for r in runs)
                if runs and e["old"] in full:
                    hits += 1
                    runs[0].text = full.replace(e["old"], e["new"])
                    for r in runs[1:]:
                        r._r.getparent().remove(r._r)
        if hits != 1:
            sys.exit(f"slide {e['slide']} paragraph text: expected 1 match, found {hits}: {e['old'][:80]!r}")
        para_touched.add(str(slide.part.partname).lstrip("/"))

    drop = sorted(set(edits.get("drop_slides", [])), reverse=True)
    if drop:
        sld_ids = prs.slides._sldIdLst
        ids = list(sld_ids)
        for n in drop:
            sid = ids[n - 1]
            rid = sid.get(f"{{{NS['r']}}}id")
            num_id = sid.get("id")
            prs.part.drop_rel(rid)
            sld_ids.remove(sid)
            # sections (p14:sectionLst) also list slide ids
            for el in prs.part._element.iter():
                if el.tag.endswith("}sldId") and el is not sid and el.get("id") == num_id and el.getparent() is not None and el.getparent() is not sld_ids:
                    el.getparent().remove(el)

    tmp = tempfile.NamedTemporaryFile(suffix=".pptx", delete=False).name
    prs.save(tmp)
    touched = {str(slides[n - 1].notes_slide.part.partname).lstrip("/") for n in changed} | para_touched
    with zipfile.ZipFile(a.deck) as zs, zipfile.ZipFile(tmp) as zt:
        src_names, tmp_names = zs.namelist(), set(zt.namelist())
        if drop:
            # slide removal changes the package structure, so keep python-pptx's copy of the structural parts
            touched |= {"ppt/presentation.xml", "ppt/_rels/presentation.xml.rels", "[Content_Types].xml", "docProps/app.xml"}
        parts, infos = {}, {}
        for info in zs.infolist():
            n = info.filename
            if n.endswith("/") or (drop and n not in tmp_names):
                continue
            parts[n] = zt.read(n) if n in touched else zs.read(n)
            infos[n] = info
    os.remove(tmp)

    slide_names = [n for n in slide_parts(zipfile.ZipFile(a.deck))]
    for e in edits.get("slide_text", []):
        part = slide_names[e["slide"] - 1]
        old = e["old"].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").encode()
        new = e["new"].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").encode()
        count = parts[part].count(old)
        if count != 1:
            sys.exit(f"slide {e['slide']} text: expected 1 match, found {count}: {e['old'][:80]!r}")
        parts[part] = parts[part].replace(old, new)

    removed = _strip_comments(parts)
    if "docProps/core.xml" in parts and edits.get("core"):
        parts["docProps/core.xml"] = _set_core(parts["docProps/core.xml"], edits["core"])
    if "docProps/app.xml" in parts:
        parts["docProps/app.xml"] = _clear_app(parts["docProps/app.xml"])

    with zipfile.ZipFile(a.out, "w", zipfile.ZIP_DEFLATED) as zo:
        order = [n for n in src_names if n in parts]
        for n in order:
            info = infos[n]
            zi = zipfile.ZipInfo(n, date_time=info.date_time)
            zi.compress_type = info.compress_type
            zo.writestr(zi, parts[n])
    print(json.dumps(dict(notes_changed=sorted(changed), dropped=sorted(drop), comments_removed=removed, out=a.out)))


def _powerpoint_running():
    r = subprocess.run(["pgrep", "-x", "Microsoft PowerPoint"], capture_output=True)
    return r.returncode == 0


def pdf_powerpoint(deck, out, timeout=600):
    """Export with PowerPoint for Mac. It is sandboxed, so work inside its container, close only our copy,
    and quit PowerPoint only if it was not running before."""
    was_running = _powerpoint_running()
    if not was_running:
        # A cold start can take longer than the open-and-save script allows, so wait until PowerPoint answers.
        subprocess.run(["open", "-g", "-a", "Microsoft PowerPoint"], capture_output=True)
        for _ in range(60):
            r = subprocess.run(["osascript", "-e", 'tell application "Microsoft PowerPoint" to count of presentations'],
                               capture_output=True, text=True)
            if r.returncode == 0:
                break
            time.sleep(1)
        time.sleep(3)
    c = os.path.expanduser("~/Library/Containers/com.microsoft.Powerpoint/Data/Documents")
    os.makedirs(c, exist_ok=True)
    tag = f"pubexport-{os.getpid()}-{int(time.time())}"
    src, dst = os.path.join(c, tag + ".pptx"), os.path.join(c, tag + ".pdf")
    shutil.copyfile(deck, src)
    script = f'''
with timeout of {timeout} seconds
  tell application "Microsoft PowerPoint"
    open POSIX file "{src}"
    delay 2
    set thePres to missing value
    repeat with p in (get presentations)
      if name of p is "{tag}.pptx" then set thePres to p
    end repeat
    save thePres in POSIX file "{dst}" as save as PDF
    close thePres saving no
  end tell
end timeout'''
    try:
        r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=timeout + 30)
        if r.returncode != 0:
            raise RuntimeError(r.stderr.strip())
        prev = -1
        for _ in range(120):
            s = os.path.getsize(dst) if os.path.exists(dst) else 0
            if s > 0 and s == prev:
                break
            prev = s
            time.sleep(2)
        if not os.path.exists(dst) or os.path.getsize(dst) == 0:
            raise RuntimeError("PowerPoint wrote no PDF")
        shutil.move(dst, out)
    finally:
        cleanup = f'''
with timeout of 60 seconds
  tell application "Microsoft PowerPoint"
    repeat with p in (get presentations)
      if name of p is "{tag}.pptx" then close p saving no
    end repeat
    {"" if was_running else "if (count of presentations) is 0 then quit saving no"}
  end tell
end timeout'''
        subprocess.run(["osascript", "-e", cleanup], capture_output=True, text=True)
        for f in (src, dst):
            if os.path.exists(f):
                os.remove(f)


def pdf_libreoffice(deck, out, dpi=200):
    work = tempfile.mkdtemp(prefix="deckpdf-")
    try:
        conf = os.path.join(work, "fonts.conf")
        dirs = ["/Applications/LibreOffice.app/Contents/Resources/fonts", "/System/Library/Fonts", "/Library/Fonts",
                os.path.expanduser("~/Library/Fonts"), "/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts"]
        with open(conf, "w") as f:
            f.write('<?xml version="1.0"?><!DOCTYPE fontconfig SYSTEM "fonts.dtd"><fontconfig>')
            f.write("".join(f"<dir>{d}</dir>" for d in dirs if os.path.isdir(d)))
            f.write(f"<cachedir>{work}/fccache</cachedir></fontconfig>")
        shutil.copyfile(deck, os.path.join(work, "deck.pptx"))
        opts = json.dumps({"ReduceImageResolution": {"type": "boolean", "value": "true"},
                           "MaxImageResolution": {"type": "long", "value": str(dpi)},
                           "Quality": {"type": "long", "value": "85"},
                           "ExportNotesPages": {"type": "boolean", "value": "false"}})
        env = dict(os.environ, FONTCONFIG_FILE=conf)
        subprocess.run(["soffice", f"-env:UserInstallation=file://{work}/profile", "--headless",
                        "--convert-to", f"pdf:impress_pdf_Export:{opts}", "deck.pptx"],
                       cwd=work, env=env, capture_output=True, timeout=900, check=True)
        shutil.move(os.path.join(work, "deck.pdf"), out)
    finally:
        shutil.rmtree(work, ignore_errors=True)


def shrink_pdf(path, max_mb):
    import fitz

    if os.path.getsize(path) <= max_mb * 1e6:
        return False
    tmp = path + ".tmp.pdf"
    for dpi, q in ((150, 80), (110, 70), (90, 60)):
        doc = fitz.open(path)
        doc.rewrite_images(dpi_threshold=dpi + 10, dpi_target=dpi, quality=q, lossy=True, lossless=True, bitonal=True, color=True, gray=True)
        doc.save(tmp, garbage=4, deflate=True)
        doc.close()
        if os.path.getsize(tmp) <= max_mb * 1e6:
            break
    os.replace(tmp, path)
    return True


def pdf_pages(path):
    out = subprocess.run(["pdfinfo", path], capture_output=True, text=True).stdout
    m = re.search(r"^Pages:\s+(\d+)", out, re.M)
    return int(m.group(1)) if m else 0


def cmd_pdf(a):
    engine = a.engine
    used = None
    if engine in ("auto", "powerpoint"):
        for attempt in (1, 2):
            try:
                pdf_powerpoint(a.deck, a.out)
                used = "powerpoint"
                break
            except Exception as ex:
                if attempt == 2:
                    if engine == "powerpoint":
                        sys.exit(f"PowerPoint export failed: {ex}")
                    print(f"PowerPoint export failed twice ({ex}); using LibreOffice", file=sys.stderr)
                else:
                    time.sleep(5)
    if used is None:
        pdf_libreoffice(a.deck, a.out)
        used = "libreoffice"
    before = os.path.getsize(a.out)
    shrunk = shrink_pdf(a.out, a.max_mb)
    print(json.dumps(dict(engine=used, pages=pdf_pages(a.out), mb=round(os.path.getsize(a.out) / 1e6, 1),
                          shrunk_from_mb=round(before / 1e6, 1) if shrunk else None, out=a.out)))


def cmd_thumb(a):
    base = tempfile.mktemp(prefix="thumb-")
    subprocess.run(["pdftoppm", "-f", "1", "-l", "1", "-singlefile", "-jpeg", "-jpegopt", "quality=85",
                    "-scale-to-x", str(a.width), "-scale-to-y", "-1", a.pdf, base], check=True)
    shutil.move(base + ".jpg", a.out)
    print(json.dumps(dict(out=a.out, kb=round(os.path.getsize(a.out) / 1024))))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("inspect"); p.add_argument("deck"); p.add_argument("--out", required=True); p.add_argument("--terms")
    p = sub.add_parser("apply"); p.add_argument("deck"); p.add_argument("edits"); p.add_argument("out")
    p = sub.add_parser("pdf"); p.add_argument("deck"); p.add_argument("out")
    p.add_argument("--engine", default="auto", choices=["auto", "powerpoint", "libreoffice"]); p.add_argument("--max-mb", type=float, default=20)
    p = sub.add_parser("thumb"); p.add_argument("pdf"); p.add_argument("out"); p.add_argument("--width", type=int, default=1280)
    a = ap.parse_args()
    {"inspect": cmd_inspect, "apply": cmd_apply, "pdf": cmd_pdf, "thumb": cmd_thumb}[a.cmd](a)


if __name__ == "__main__":
    main()
