#!/usr/bin/env python3
"""Build and check the talk catalog.

  catalog.py build            Write every talk README, the root README, and talks.json from the talk.yml files.
  catalog.py check [--links]  Validate talk.yml files, generated files, file sizes, and blocked content.
                              --links also checks that each URL resolves and each release asset exists.
  catalog.py new --date 2026-10-27 --event "Techorama NL 2026" --title "..." [--slug short-name]
                              Start a talk folder with a talk.yml to fill in.
  catalog.py release SLUG DECK.pptx
                              Create the GitHub release for a talk and upload its public deck. Run it after
                              the talk folder is pushed, because the release tag points at the pushed commit.

Each talk lives in YEAR/YYYY-MM-DD-slug/ with talk.yml (the source of truth), README.md (generated),
slides.pdf and thumbnail.jpg. The editable deck is a release asset, never a file in git.
"""
import argparse
import datetime as dt
import glob
import json
import os
import re
import subprocess
import sys
import urllib.error
import urllib.parse
import urllib.request

import yaml

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OWNER, REPO = "davidortinau", "presentations"
REPO_URL = f"https://github.com/{OWNER}/{REPO}"
FORMATS = {"talk", "keynote", "workshop", "webinar", "panel", "livestream", "session"}
RECORDED = {"yes", "no", "unknown"}
NOTES = {"reviewed", "removed", "none"}
PDF_WARN_MB, PDF_MAX_MB, FILE_MAX_MB = 10, 20, 50
BLOCKED = [
    (r"davidos", "link or reference to the private davidos repo"),
    (r"\.\./\.\./(reports|context|journal|Projects)/", "relative link into a private repo"),
    (r"/Users/[A-Za-z]", "local file path"),
    (r"sharepoint\.com|teams\.microsoft\.com|dev\.azure\.com|\beng\.ms\b|microsoft-my\.", "internal Microsoft host"),
    (r"(?i)\bmicrosoft confidential\b|\binternal only\b", "confidentiality marking"),
]
INTRO = """# Talks by David Ortinau

Slides, abstracts, recordings, and demo code from my conference talks, webinars, and community sessions. I'm a Principal Product Manager for .NET at Microsoft, and my team works on AI libraries and experiences, Aspire, ASP.NET Core and Blazor, and .NET MAUI.

Each talk folder has the abstract, a PDF of the slides, the PowerPoint with speaker notes when I've shared it, the recording when there is one, and the demo code as presented. Newest first.
"""
RIGHTS = """## About these files

The slides and notes are shared for reference. They are not offered under an open-source license. Demo code lives in each demo repository under that repository's license. Product names and logos belong to their owners.

This page is generated from each talk's `talk.yml` by [`tools/catalog.py`](tools/catalog.py). To add a talk, run `python3 tools/catalog.py new`, fill in the `talk.yml`, then run `python3 tools/catalog.py build` and `python3 tools/catalog.py check`.
"""


def load_talks():
    talks = []
    for path in sorted(glob.glob(os.path.join(ROOT, "[0-9][0-9][0-9][0-9]", "*", "talk.yml"))):
        with open(path, encoding="utf-8") as f:
            t = yaml.safe_load(f) or {}
        t["_dir"] = os.path.dirname(path)
        t["_slug"] = os.path.basename(t["_dir"])
        t["_year"] = os.path.basename(os.path.dirname(t["_dir"]))
        t["_rel"] = os.path.relpath(t["_dir"], ROOT).replace(os.sep, "/")
        talks.append(t)
    talks.sort(key=lambda t: (str(t.get("date")), t["_slug"]), reverse=True)
    return talks


def load_list(name):
    path = os.path.join(ROOT, name)
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f) or []


class TalkDate:
    """A talk date as precise as the sources allow: YYYY-MM-DD, YYYY-MM, or YYYY."""

    def __init__(self, v):
        text = v.isoformat() if isinstance(v, dt.date) else str(v)
        if not re.fullmatch(r"\d{4}(-\d{2}(-\d{2})?)?", text):
            raise ValueError(text)
        parts = [int(x) for x in text.split("-")]
        self.text = text
        self.precision = ("year", "month", "day")[len(parts) - 1]
        self.date = dt.date(parts[0], parts[1] if len(parts) > 1 else 1, parts[2] if len(parts) > 2 else 1)
        self.year = self.date.year

    def isoformat(self):
        return self.text

    def __str__(self):
        return self.text


def as_date(v):
    return v if isinstance(v, TalkDate) else TalkDate(v)


def long_date(d):
    d = as_date(d)
    if d.precision == "year":
        return str(d.year)
    if d.precision == "month":
        return d.date.strftime("%B %Y")
    return f"{d.date.strftime('%A')} {d.date.day} {d.date.strftime('%B %Y')}"


def short_date(d):
    d = as_date(d)
    if d.precision == "year":
        return ""
    if d.precision == "month":
        return d.date.strftime("%b")
    return f"{d.date.day} {d.date.strftime('%b')}"


def release_url(t):
    r = (t.get("slides") or {}).get("release")
    if not r:
        return None
    return f"{REPO_URL}/releases/download/{r['tag']}/{urllib.parse.quote(r['asset'])}"


def gh_tree(repo, ref):
    return f"https://github.com/{repo}/tree/{urllib.parse.quote(ref, safe='/')}"


def gh_blob(repo, ref, path):
    return f"https://github.com/{repo}/blob/{urllib.parse.quote(ref, safe='/')}/{path}"


def demo_url(dm):
    ref = dm.get("commit") or dm.get("branch")
    return gh_tree(dm["repo"], ref) if ref else f"https://github.com/{dm['repo']}"


def md_escape(s):
    return str(s).replace("|", "\\|")


def video_links(t):
    return [v for v in (t.get("video") or []) if v.get("url")]


def render_talk(t, by_slug):
    d = as_date(t["date"])
    slides = t.get("slides") or {}
    L = [f"# {t['title']}", ""]
    where = [f"**{t['event']}**"]
    place = ", ".join(x for x in (t.get("venue"), t.get("city")) if x)
    if place:
        where.append(place)
    when = long_date(d) + (f", {t['time']}" if t.get("time") else "")
    where.append(when)
    if t.get("length_min"):
        where.append(f"{t['length_min']} minutes")
    if t.get("session_url"):
        where.append(f"[Session page]({t['session_url']})")
    L.append(" · ".join(where))
    if t.get("co_speakers"):
        names = []
        for c in t["co_speakers"]:
            names.append(f"[{c['name']}](https://github.com/{c['github']})" if c.get("github") else c["name"])
        L += ["", "With " + ", ".join(names) + "."]
    if t.get("note"):
        L += ["", t["note"].strip()]
    if os.path.exists(os.path.join(t["_dir"], "thumbnail.jpg")):
        target = slides.get("pdf") or "thumbnail.jpg"
        L += ["", f"[![First slide of {t['title']}](thumbnail.jpg)]({target})"]

    L += ["", "## Abstract", ""]
    if t.get("abstract"):
        for para in t["abstract"].strip().split("\n\n"):
            L.append("> " + para.strip().replace("\n", "\n> "))
            L.append(">")
        L.pop()
        src = t.get("abstract_source")
        if src:
            L += ["", f"From the [{t.get('abstract_source_label', 'session page')}]({src})."]
    else:
        L.append("No published abstract.")

    L += ["", "## Slides", ""]
    items = []
    if slides.get("pdf"):
        pdf_path = os.path.join(t["_dir"], slides["pdf"])
        mb = os.path.getsize(pdf_path) / 1e6 if os.path.exists(pdf_path) else 0
        pages = f"{slides['pages']} slides, " if slides.get("pages") else ""
        items.append(f"- [PDF]({slides['pdf']}) ({pages}{mb:.1f} MB)")
    url = release_url(t)
    if url:
        r = slides["release"]
        note = {"reviewed": " The speaker notes were reviewed for public sharing.",
                "removed": " The speaker notes were removed.",
                "none": " This deck has no speaker notes."}.get(slides.get("notes"), "")
        size = f" ({r['mb']:.1f} MB)" if r.get("mb") else ""
        items.append(f"- [PowerPoint]({url}){size}.{note}")
    for a in slides.get("also_at") or []:
        items.append(f"- [{a['label']}]({a['url']})")
    L += items or ["Not shared."]

    L += ["", "## Video", ""]
    vids = video_links(t)
    if vids:
        for v in vids:
            L.append(f"- [{v.get('label', 'Recording')}]({v['url']})")
    else:
        L.append({"no": "Not recorded.", "unknown": "No recording posted yet."}.get(t.get("recorded", "unknown"), "No recording posted yet."))

    demos = t.get("demos") or []
    if demos:
        L += ["", "## Demos", "", "| Demo | Code | As presented |", "|---|---|---|"]
        for dm in demos:
            repo = dm["repo"]
            code = f"[{repo}]({gh_tree(repo, dm['branch'])}), branch `{dm['branch']}`" if dm.get("branch") else f"[{repo}](https://github.com/{repo})"
            pres = []
            if dm.get("commit"):
                pres.append(f"[{dm['commit'][:7]}]({gh_tree(repo, dm['commit'])})")
            for doc in dm.get("docs") or []:
                ref = dm.get("commit") or dm.get("branch") or "HEAD"
                pres.append(f"[{doc['label']}]({gh_blob(repo, ref, doc['path'])})")
            if dm.get("note"):
                pres.append(dm["note"])
            L.append(f"| {md_escape(dm['name'])} | {code} | {' · '.join(pres) or ' '} |")

    same = [s for s in (t.get("same_talk_as") or []) if s in by_slug]
    if same:
        L += ["", "## Also presented at", ""]
        for s in same:
            o = by_slug[s]
            L.append(f"- [{o['event']}, {long_date(as_date(o['date']))}](../../{o['_rel']}/)")

    if t.get("resources"):
        L += ["", "## Resources", ""]
        for r in t["resources"]:
            L.append(f"- [{r['label']}]({r['url']})")

    L += ["", "---", "", "[All talks](../../README.md)", ""]
    return "\n".join(L)


def render_root(talks):
    L = [INTRO]
    years = {}
    for t in talks:
        years.setdefault(t["_year"], []).append(t)
    for year in sorted(years, reverse=True):
        L += [f"## {year}", "", "| Date | Talk | Event | Slides | Video | Demos |", "|---|---|---|---|---|---|"]
        for t in years[year]:
            d = as_date(t["date"])
            slides = t.get("slides") or {}
            s = []
            if slides.get("pdf"):
                s.append(f"[PDF]({t['_rel']}/{slides['pdf']})")
            if release_url(t):
                s.append(f"[PPTX]({release_url(t)})")
            v = video_links(t)
            vid = f"[Watch]({v[0]['url']})" if v else ""
            demos = " · ".join(f"[{dm['repo'].split('/')[-1]}]({demo_url(dm)})" for dm in (t.get("demos") or []))
            event = t["event"] + (f", {t['city'].split(',')[0]}" if t.get("city") and t["city"].lower() != "online" else "")
            L.append(f"| {short_date(d)} | [{md_escape(t['title'])}]({t['_rel']}/) | {md_escape(event)} | {' · '.join(s)} | {vid} | {demos} |")
        L.append("")
    series = load_list("series.yml")
    if series:
        L += ["## Shows and series", "", "| Show | Years | Role | Link |", "|---|---|---|---|"]
        for s in series:
            L.append(f"| {md_escape(s['name'])} | {s.get('years', '')} | {md_escape(s.get('role', ''))} | [{s.get('link_label', 'Watch')}]({s['url']}) |")
        L.append("")
    pods = load_list("podcasts.yml")
    if pods:
        L += ["## Podcasts and interviews", "", "| Date | Show | Episode |", "|---|---|---|"]
        for p in sorted(pods, key=lambda p: str(p.get("date", "")), reverse=True):
            L.append(f"| {p.get('date', '')} | {md_escape(p['show'])} | [{md_escape(p.get('title', 'Listen'))}]({p['url']}) |")
        L.append("")
    L.append(RIGHTS)
    return "\n".join(L)


def talks_json(talks):
    out = []
    for t in talks:
        slides = t.get("slides") or {}
        thumb = os.path.join(t["_dir"], "thumbnail.jpg")
        out.append({
            "slug": t["_slug"],
            "title": t["title"],
            "date": str(as_date(t["date"])),
            "event": t["event"],
            "city": t.get("city"),
            "format": t.get("format"),
            "url": f"{REPO_URL}/tree/main/{t['_rel']}",
            "slides_pdf": f"{REPO_URL}/blob/main/{t['_rel']}/{slides['pdf']}" if slides.get("pdf") else None,
            "slides_pptx": release_url(t),
            "thumbnail": f"https://raw.githubusercontent.com/{OWNER}/{REPO}/main/{t['_rel']}/thumbnail.jpg" if os.path.exists(thumb) else None,
            "video": [v["url"] for v in video_links(t)],
            "tags": t.get("tags") or [],
        })
    return json.dumps({"talks": out}, indent=2, ensure_ascii=False) + "\n"


def generated(talks):
    by_slug = {t["_slug"]: t for t in talks}
    files = {os.path.join(t["_dir"], "README.md"): render_talk(t, by_slug) for t in talks}
    files[os.path.join(ROOT, "README.md")] = render_root(talks)
    files[os.path.join(ROOT, "talks.json")] = talks_json(talks)
    return files


def validate(t):
    errs = []
    where = t["_rel"]
    for k in ("title", "date", "event", "format"):
        if not t.get(k):
            errs.append(f"{where}: missing {k}")
    if errs:
        return errs
    try:
        d = as_date(t["date"])
    except ValueError:
        return [f"{where}: date must be YYYY-MM-DD, YYYY-MM, or YYYY"]
    if not re.match(re.escape(d.isoformat()) + r"-[a-z]", t["_slug"]):
        errs.append(f"{where}: folder must start with the date {d.isoformat()}")
    if t["_year"] != str(d.year):
        errs.append(f"{where}: folder must sit under {d.year}/")
    if not re.fullmatch(r"\d{4}(-\d{2}(-\d{2})?)?-[a-z0-9]+(-[a-z0-9]+)*", t["_slug"]):
        errs.append(f"{where}: folder name must be lowercase words joined by hyphens")
    if t["format"] not in FORMATS:
        errs.append(f"{where}: format must be one of {sorted(FORMATS)}")
    if t.get("recorded", "unknown") not in RECORDED:
        errs.append(f"{where}: recorded must be one of {sorted(RECORDED)}")
    slides = t.get("slides") or {}
    if os.path.exists(os.path.join(t["_dir"], "slides.pdf")) and slides.get("pdf") != "slides.pdf":
        errs.append(f"{where}: slides.pdf is in the folder but talk.yml does not list it")
    if slides.get("pdf") and not os.path.exists(os.path.join(t["_dir"], slides["pdf"])):
        errs.append(f"{where}: {slides['pdf']} is missing")
    if slides.get("notes") and slides["notes"] not in NOTES:
        errs.append(f"{where}: slides.notes must be one of {sorted(NOTES)}")
    r = slides.get("release")
    if r and not (r.get("tag") and r.get("asset")):
        errs.append(f"{where}: slides.release needs tag and asset")
    if r and slides.get("notes") not in NOTES:
        errs.append(f"{where}: a shared deck needs slides.notes (reviewed, removed or none)")
    if t.get("abstract") and not t.get("abstract_source"):
        errs.append(f"{where}: abstract needs abstract_source")
    for dm in t.get("demos") or []:
        if not dm.get("repo") or "/" not in dm["repo"] or not dm.get("name"):
            errs.append(f"{where}: each demo needs name and repo (owner/name)")
        if dm.get("commit") and not re.fullmatch(r"[0-9a-f]{40}", dm["commit"]):
            errs.append(f"{where}: demo commit must be a full 40-character SHA")
    for v in t.get("video") or []:
        if not v.get("url"):
            errs.append(f"{where}: each video needs a url")
    return errs


def tracked_files():
    out = subprocess.run(["git", "ls-files", "-z", "--cached", "--others", "--exclude-standard"], cwd=ROOT,
                         capture_output=True, text=True, check=True).stdout
    return [f for f in out.split("\0") if f]


def urls_of(t):
    urls = [t.get("session_url"), t.get("abstract_source"), t.get("event_url")]
    urls += [v.get("url") for v in t.get("video") or []]
    urls += [r.get("url") for r in t.get("resources") or []]
    urls += [a.get("url") for a in (t.get("slides") or {}).get("also_at") or []]
    for dm in t.get("demos") or []:
        if dm.get("commit"):
            urls.append(gh_tree(dm["repo"], dm["commit"]))
        elif dm.get("branch"):
            urls.append(gh_tree(dm["repo"], dm["branch"]))
        else:
            urls.append(f"https://github.com/{dm['repo']}")
        for doc in dm.get("docs") or []:
            urls.append(gh_blob(dm["repo"], dm.get("commit") or dm.get("branch") or "HEAD", doc["path"]))
    return [u for u in urls if u]


def url_ok(url):
    """True when the URL resolves, False when it is broken, None when the site blocks automated checks."""
    if re.search(r"youtube\.com/watch|youtu\.be/", url):
        url = "https://www.youtube.com/oembed?format=json&url=" + urllib.parse.quote(url, safe="")
    for method in ("HEAD", "GET"):
        req = urllib.request.Request(url, method=method, headers={"User-Agent": "Mozilla/5.0 (talk-catalog link check)"})
        try:
            with urllib.request.urlopen(req, timeout=20) as r:
                return r.status < 400, r.status
        except urllib.error.HTTPError as e:
            if method == "HEAD" and e.code in (401, 403, 405, 429, 999):
                continue
            if e.code in (401, 403, 429, 999):
                return None, e.code
            return False, e.code
        except Exception as e:
            if method == "HEAD":
                continue
            return False, type(e).__name__
    return False, "?"


def release_assets(tag):
    r = subprocess.run(["gh", "release", "view", tag, "--repo", f"{OWNER}/{REPO}", "--json", "assets"],
                       capture_output=True, text=True)
    if r.returncode != 0:
        return None
    return [a["name"] for a in json.loads(r.stdout)["assets"]]


def cmd_build(a):
    talks = load_talks()
    errs = [e for t in talks for e in validate(t)]
    if errs:
        print("\n".join(errs))
        sys.exit(1)
    for path, text in generated(talks).items():
        old = open(path, encoding="utf-8").read() if os.path.exists(path) else None
        if old != text:
            with open(path, "w", encoding="utf-8") as f:
                f.write(text)
            print("wrote", os.path.relpath(path, ROOT))
    print(f"{len(talks)} talks")


def cmd_check(a):
    talks = load_talks()
    errs, warns = [], []
    for t in talks:
        errs += validate(t)
    slugs = {t["_slug"] for t in talks}
    for t in talks:
        for s in t.get("same_talk_as") or []:
            if s not in slugs:
                errs.append(f"{t['_rel']}: same_talk_as {s} is not a talk folder")
    if not errs:
        for path, text in generated(talks).items():
            cur = open(path, encoding="utf-8").read() if os.path.exists(path) else None
            if cur != text:
                errs.append(f"{os.path.relpath(path, ROOT)} is out of date; run tools/catalog.py build")
    for f in tracked_files():
        full = os.path.join(ROOT, f)
        if not os.path.exists(full):
            continue
        if re.search(r"\.(pptx|ppt|key|potx)$", f, re.I):
            errs.append(f"{f}: decks go in a GitHub release, not in git")
        mb = os.path.getsize(full) / 1e6
        if mb > FILE_MAX_MB:
            errs.append(f"{f}: {mb:.1f} MB is over {FILE_MAX_MB} MB")
        if f.lower().endswith(".pdf"):
            if mb > PDF_MAX_MB:
                errs.append(f"{f}: {mb:.1f} MB is over {PDF_MAX_MB} MB; run tools/deck.py pdf to shrink it")
            elif mb > PDF_WARN_MB:
                warns.append(f"{f}: {mb:.1f} MB")
        if re.search(r"\.(md|yml|yaml|json|txt)$", f) and not f.startswith("tools/") and not f.startswith(".github/"):
            text = open(full, encoding="utf-8", errors="replace").read()
            for pat, why in BLOCKED:
                m = re.search(pat, text)
                if m:
                    errs.append(f"{f}: {why} ({m.group(0)!r})")
    if a.links:
        for t in talks:
            for u in sorted(set(urls_of(t))):
                ok, status = url_ok(u)
                if ok is None:
                    warns.append(f"{t['_rel']}: {u} blocks automated checks ({status}); open it by hand")
                elif not ok:
                    errs.append(f"{t['_rel']}: {u} returned {status}")
            r = (t.get("slides") or {}).get("release")
            if r:
                assets = release_assets(r["tag"])
                if assets is None:
                    errs.append(f"{t['_rel']}: release {r['tag']} does not exist")
                elif r["asset"] not in assets:
                    errs.append(f"{t['_rel']}: release {r['tag']} has no asset {r['asset']}")
    for w in warns:
        print("warning:", w)
    for e in errs:
        print("error:", e)
    print(f"{len(talks)} talks, {len(errs)} errors, {len(warns)} warnings")
    sys.exit(1 if errs else 0)


def slugify(s):
    s = re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")
    return re.sub(r"-+", "-", s)


def cmd_new(a):
    d = as_date(a.date)
    slug = f"{d.isoformat()}-{slugify(a.slug or (a.event + ' ' + a.title))}"
    folder = os.path.join(ROOT, str(d.year), slug)
    if os.path.exists(folder):
        sys.exit(f"{os.path.relpath(folder, ROOT)} already exists")
    os.makedirs(folder)
    tmpl = f"""title: {json.dumps(a.title)}
date: {d.isoformat()}
time:                 # "16:00", local time at the event
length_min:
event: {json.dumps(a.event)}
event_url:
venue:
city:                 # "Utrecht, Netherlands", or "Online"
format: talk          # talk, keynote, workshop, webinar, panel, livestream, session
session_url:
abstract: |
  Paste the published abstract verbatim.
abstract_source:      # the page the abstract came from
co_speakers: []       # - {{name: Maddy Montaquila, github: maddymontaquila}}
slides:
  pdf: slides.pdf
  pages:
  notes: reviewed     # reviewed, removed, none
  release:
    tag: {slug}
    asset: {slug}.pptx
    mb:
recorded: unknown     # yes, no, unknown
video: []             # - {{label: Recording, url: https://www.youtube.com/watch?v=...}}
demos: []             # - {{name: ..., repo: owner/name, branch: ..., commit: <40-char sha>, docs: [{{label: Run sheet, path: ...}}]}}
same_talk_as: []      # other talk folder names
resources: []         # - {{label: ..., url: ...}}
tags: []
sources: []           # where each fact came from
"""
    with open(os.path.join(folder, "talk.yml"), "w", encoding="utf-8") as f:
        f.write(tmpl)
    print(os.path.relpath(folder, ROOT))


def cmd_release(a):
    talks = {t["_slug"]: t for t in load_talks()}
    t = talks.get(a.slug)
    if not t:
        sys.exit(f"no talk folder named {a.slug}")
    r = (t.get("slides") or {}).get("release")
    if not r:
        sys.exit("talk.yml has no slides.release")
    if os.path.basename(a.deck) != r["asset"]:
        sys.exit(f"the deck file must be named {r['asset']}")
    notes = f"{t['title']}, {t['event']}, {long_date(as_date(t['date']))}.\n\nAbstract, PDF, video and demos: {REPO_URL}/tree/main/{t['_rel']}"
    if release_assets(r["tag"]) is None:
        cmd = ["gh", "release", "create", r["tag"], a.deck, "--repo", f"{OWNER}/{REPO}", "--title", f"{t['title']} ({t['event']})", "--notes", notes, "--target", "main"]
    else:
        cmd = ["gh", "release", "upload", r["tag"], a.deck, "--repo", f"{OWNER}/{REPO}", "--clobber"]
    subprocess.run(cmd, check=True)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("build")
    p = sub.add_parser("check"); p.add_argument("--links", action="store_true")
    p = sub.add_parser("new"); p.add_argument("--date", required=True); p.add_argument("--event", required=True)
    p.add_argument("--title", required=True); p.add_argument("--slug")
    p = sub.add_parser("release"); p.add_argument("slug"); p.add_argument("deck")
    a = ap.parse_args()
    {"build": cmd_build, "check": cmd_check, "new": cmd_new, "release": cmd_release}[a.cmd](a)


if __name__ == "__main__":
    main()
