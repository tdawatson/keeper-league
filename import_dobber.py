#!/usr/bin/env python3
"""Scrape DobberProspects for the ratings that matter to a keeper league.

Three fields per player, off his public profile page: Fantasy Upside and NHL
Certainty (both 0-10, from the Attributes table) and the one-line Fantasy
Summary. Example: dobberprospects.com/player/chase-reid/ -> 9.0, 9.0,
"Smooth-skating two-way defender with strong offensive instincts."

Profile URLs come from the site's own player sitemap rather than being guessed
from the name - spellings differ (Egor/Yegor Shilov, Artyon/Artyom Matyuk), so
an exact match on the normalised slug is tried first and a close match second,
with ALIASES for the ones no amount of string-matching will reach.

Whatever is left over really is absent. Every unmatched name was put through
the site's own player search (/wp-json/wp/v2/player?search=) on 2026-09-19: 32
of the 33 returned no profile under any spelling, and all 32 are picks 85 or
later, which is simply past the depth DobberProspects writes up. They are
written out with blank ratings rather than dropped, so every pick keeps a row.

Two modes. With no arguments it does the entry-draft class from nhl_draft.py
and writes data/dobber.csv, keyed by draft pick. With --players FILE it reads
any CSV carrying a name column and writes the same three fields keyed by name,
which is how the prospect board gets its ratings:

    ./import_dobber.py --players data/prospect-board.csv \
                       --out data/dobber-prospects.csv

Pages are cached under data/dobber-cache/, so a rerun is free and --refresh is
what goes back to the network.
"""
import argparse
import csv
import difflib
import html
import pathlib
import re
import sys
import time
import urllib.error
import urllib.request

from names import norm, same_person

SITEMAP = "https://dobberprospects.com/wp-sitemap-posts-player-1.xml"
UA = {"User-Agent": "Mozilla/5.0 (keeper-league workbook; personal use)"}
CACHE = pathlib.Path("data/dobber-cache")
OUT = pathlib.Path("data/dobber.csv")
FIELDS = ["pick", "player", "upside", "certainty", "summary", "url"]

# Players the site files under a different name than the draft board does.
# Checked against the "Drafted" line on the profile, not just the surname:
# Mikey Berchild's page reads "2026 Round 4; Overall: 105", which is our pick.
ALIASES = {"Michael Berchild": "mikey-berchild"}


def get(url, refresh=False, delay=1.0):
    """Fetch a page, through the on-disk cache. Returns None on any failure."""
    CACHE.mkdir(parents=True, exist_ok=True)
    f = CACHE / (re.sub(r"[^A-Za-z0-9]+", "-", url).strip("-") + ".html")
    if f.exists() and not refresh:
        return f.read_text(encoding="utf-8")
    try:
        with urllib.request.urlopen(
                urllib.request.Request(url, headers=UA), timeout=30) as r:
            body = r.read().decode("utf-8", "replace")
    except (urllib.error.URLError, OSError) as e:
        print(f"  ! {url}: {e}")
        return None
    f.write_text(body, encoding="utf-8")
    time.sleep(delay)                   # be a polite guest on their server
    return body


def text(s):
    """HTML fragment -> flat text."""
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", html.unescape(s)).strip()


def attribute(page, label):
    """One row of the Attributes table, looked up by its row header.

    The markup omits closing </th>/</td> tags, so the value runs from the end
    of the header cell to whatever opens next."""
    m = re.search(r'<th scope="row">\s*' + label + r'\b.*?<td>(.*?)(?=<tr|</t)',
                  page, re.S | re.I)
    return text(m.group(1)) if m else ""


def summary(page):
    """The Fantasy Summary paragraph - one sentence, under its own banner."""
    m = re.search(r'headerBand[^>]*>\s*Fantasy Summary\s*</h\d>\s*<p>(.*?)</p>',
                  page, re.S | re.I)
    return text(m.group(1)) if m else ""


BORN = re.compile(r"\b(19|20)\d{2}\b")


def born(page):
    """Birth year off the profile's Date of Birth row, if it has a usable one.

    A year that would make the player an infant or a pensioner is the site's
    own typo - Ludvig Andersson's page reads "May 24, 2026" - and says
    nothing about whether this is the right man, so it is discarded rather
    than counted as a mismatch."""
    m = BORN.search(attribute(page, "Date of Birth"))
    if not m:
        return None
    year = int(m.group(0))
    return year if 14 <= SEASON_YEAR - year <= 60 else None


def slug_index(refresh, delay):
    """Every player URL the site publishes: normalised slug -> (slug, url)."""
    xml = get(SITEMAP, refresh, delay)
    if not xml:
        sys.exit("could not read the player sitemap")
    urls = re.findall(r"<loc>(https://dobberprospects\.com/player/[^/]+/)</loc>",
                      xml)
    return {norm(u.rsplit("/", 2)[-2]): (u.rsplit("/", 2)[-2], u) for u in urls}


def resolve(name, index, keys):
    """A name -> profile URL. Alias, then exact, then a close spelling.

    "Close" has to hold for BOTH names. Scoring the whole slug in one go
    quietly paired Dylan Roobroeck with his brother Ryan, Jack O'Brien with
    Jake, and Michael Hagens with Michael Hage. So the surname must match
    almost exactly, and the first name must be the same name in another form -
    Zach/Zachary, Gabe/Gabriel, Andrei/Andrey, Scherbakov/Shcherbakov."""
    want = ALIASES.get(name, name)
    k = norm(want)
    if k in index:
        return index[k][1]
    for cand in difflib.get_close_matches(k, keys, n=8, cutoff=0.80):
        slug, url = index[cand]
        if same_person(want, slug):
            return url
    return None


SEASON_YEAR = 2026          # ages in the export are as of this season


def scrape(names, index, keys, refresh, delay, ages=None):
    """[(key, name)] -> a row per name, whether or not the site covers him.

    A matched profile still has to describe someone of the right age. Surnames
    one letter apart are the last thing string comparison cannot settle -
    Eric Neilson is 42 and eric-nilson was born in 2007 - and the profile's
    own date of birth settles it where a name never could."""
    ages = ages or {}
    rows, no_page, no_data, wrong_age = [], [], [], []
    for key, name in names:
        url = resolve(name, index, keys)
        if not url:
            no_page.append(name)
            rows.append({"player": name, "upside": "", "certainty": "",
                         "summary": "", "url": "", "_key": key})
            continue
        page = get(url, refresh, delay) or ""
        by = born(page)
        if by is not None and ages.get(name) is not None \
                and abs((SEASON_YEAR - by) - ages[name]) > 2:
            # the page describes someone born years apart from this player:
            # Eric Neilson is 42, eric-nilson was born in 2007
            wrong_age.append((name, ages[name], by))
            rows.append({"player": name, "upside": "", "certainty": "",
                         "summary": "", "url": "", "_key": key})
            continue
        up, cert = attribute(page, "Fantasy Upside"), attribute(page, "NHL Certainty")
        summ = summary(page)
        if not (up or cert or summ):
            no_data.append(name)
        rows.append({"player": name, "upside": up, "certainty": cert,
                     "summary": summ, "url": url, "_key": key})
    return rows, no_page, no_data, wrong_age


def dedupe(rows):
    """One profile, one player.

    Two names sometimes reach the same page - Ryan and Bryan Cameron, Ludwig
    and Ludvig Persson - and at most one of them is the man it describes.
    The closer spelling keeps it and the other is left unrated, because a
    blank cell is honest where someone else's scouting report is not."""
    from names import norm, ratio
    claims = {}
    for r in rows:
        if r["url"]:
            claims.setdefault(r["url"], []).append(r)
    stripped = []
    for url, rs in claims.items():
        if len(rs) < 2:
            continue
        slug = norm(url.rstrip("/").rsplit("/", 1)[-1])
        rs.sort(key=lambda r: (norm(r["player"]) != slug,
                               -ratio(norm(r["player"]), slug)))
        for r in rs[1:]:
            stripped.append((r["player"], rs[0]["player"]))
            r.update(upside="", certainty="", summary="", url="")
    return stripped


def report(out, rows, no_page, no_data, unit):
    rated = sum(1 for r in rows if r["upside"])
    print(f"wrote {out}  ({len(rows)} {unit}, {rated} with ratings)")
    for label, lst in (("no profile on the site", no_page),
                       ("profile but no ratings", no_data)):
        if lst:
            print(f"  {label} ({len(lst)}): {', '.join(lst[:8])}"
                  f"{' ...' if len(lst) > 8 else ''}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--players", help="CSV of players to rate instead of the "
                                      "draft class; needs a name column")
    ap.add_argument("--name-col", default="name",
                    help="which column of --players holds the name")
    ap.add_argument("--where", action="append", default=[], metavar="COL=VALUE",
                    help="only rate rows whose COL equals VALUE; repeatable")
    ap.add_argument("--refresh", action="store_true",
                    help="re-fetch every page instead of using data/dobber-cache")
    ap.add_argument("--delay", type=float, default=1.0,
                    help="seconds between requests (default 1.0)")
    ap.add_argument("--out", default=str(OUT))
    args = ap.parse_args()

    sys.path.insert(0, str(pathlib.Path(__file__).parent))
    index = slug_index(args.refresh, args.delay)
    keys = list(index)
    print(f"sitemap: {len(index)} player pages")

    ages = {}
    if args.players:
        src = pathlib.Path(args.players)
        where = [w.split("=", 1) for w in args.where]
        seen, names = set(), []
        for r in csv.DictReader(src.open()):
            if any((r.get(c) or "") != v for c, v in where):
                continue
            n = (r.get(args.name_col) or "").strip()
            if n and n not in seen:     # one page per player, however listed
                seen.add(n)
                names.append((n, n))
                try:
                    ages[n] = float(r.get("age") or "")
                except ValueError:
                    pass
        fields, unit = ["player", "upside", "certainty", "summary", "url"], "players"
    else:
        from nhl_draft import PICKS
        names = [(pick, n) for pick, _, _, n, *_ in PICKS if not n.startswith("(")]
        fields, unit = FIELDS, "picks"

    rows, no_page, no_data, wrong_age = scrape(
        names, index, keys, args.refresh, args.delay, ages)
    for nm, a, by in wrong_age:
        print(f"  ! {nm} is {a:.0f}; that profile was born {by} - left unrated")
    for loser, winner in dedupe(rows):
        print(f"  ! {loser} shared a profile with {winner}; left unrated")
    for r in rows:
        if "pick" in fields:
            r["pick"] = r.pop("_key")
        else:
            r.pop("_key")

    out = pathlib.Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("w", newline="") as f:
        w = csv.DictWriter(f, fields)
        w.writeheader()
        w.writerows(rows)
    report(out, rows, no_page, no_data, unit)


if __name__ == "__main__":
    main()
