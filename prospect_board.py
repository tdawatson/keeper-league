#!/usr/bin/env python3
"""Rank the prospect-draft board: who is still league-prospect eligible.

Rule guide s."Prospect Draft (5 rounds)": a prospect is a FWD/D who has yet to
play his 41st NHL game, or a goalie who has yet to play his 26th. That is a
CAREER count, and it is the one number the Fantrax export does not carry - the
export only knows this season - so eligibility here is settled against the NHL's
own public API rather than guessed at from 25/26 games.

That distinction decides the board. Tyson Foerster played 29 games in 25/26 and
looks eligible on the export; he has 195 career games and is not. Aleksander
Barkov played none and looks even more eligible; he has over 800.

A candidate the NHL API has never heard of has no NHL games, so he IS eligible
- that is the import bucket, the KHL/SHL/Liiga/NCAA player arriving on his first
NHL contract. The rule guide is explicit that he need not be a real-world rookie:
"A player could be 30 years old and have played 24 games and still qualify."

Writes data/prospect-board.csv. Responses cache under data/nhl-cache/.

Usage: ./prospect_board.py [--min-proj 10] [--refresh]
"""
import argparse
import concurrent.futures
import difflib
import csv
import json
import pathlib
import re
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request

from names import norm as nm_norm, same_first, same_person

SEARCH = ("https://search.d3.nhle.com/api/v1/search/player"
          "?culture=en-us&limit=200&q=")   # 50 truncates common surnames
LANDING = "https://api-web.nhle.com/v1/player/{}/landing"
UA = {"User-Agent": "Mozilla/5.0 (keeper-league workbook; personal use)"}
CACHE = pathlib.Path("data/nhl-cache")
OUT = pathlib.Path("data/prospect-board.csv")
GP_CAP = {"G": 25, "F": 40, "D": 40}
FIELDS = ["name", "team", "pos", "age", "career_gp", "gp_2526", "fpts_2526",
          "fpg_2526", "proj_2627", "eligible", "rookie", "draft_pick",
          "nhl_id", "note"]



AMBIGUOUS = object()        # surname found, but no confident first-name match


# Draftees the Fantrax export files under a different given name, so no
# amount of string comparison will reach them. Both confirmed by club:
# KJ Sauer went to PHI at 136 and the export's Kent Sauer is PHI; Noa Ta'amu
# went to WPG at 199 and the export's Alofa Tunoa Ta'amu is WPG.
FANTRAX_ALIAS = {"kentsauer": "kjsauer",
                 "alofatunoataamu": "noataamu"}
FANTRAX_ALIAS_REV = {v: k for k, v in FANTRAX_ALIAS.items()}


def assign_picks(rows):
    """Tag each board row with its 2026 entry-draft pick, one pick to one row.

    Done as a pass over the finished board rather than per player, because
    the decision needs the whole field: the pool holds a retired Bryan
    Cameron (37) and a retired Vladimir Dravecky (41) whose names sit right
    next to real 2026 draftees, and only the club tells them apart.

    A name that matches exactly and uniquely is taken on its own - Andre
    Mondoux was drafted by the Rangers and is in the Flyers organisation now,
    and a club test alone would lose him. Anything less than that, a spelling
    variant or a shared name, has to agree on club as well."""
    try:
        from nhl_draft import PICKS, TEAM_ABBR
    except ImportError:
        return
    by_name = {}
    for r in rows:
        by_name.setdefault(nm_norm(r["name"]), []).append(r)
    taken = set()

    def claim(r, pick):
        if id(r) in taken:
            return False
        r["draft_pick"] = pick
        taken.add(id(r))
        return True

    for pick, _, _, name, _, _, club, *_ in PICKS:
        if name.startswith("("):
            continue
        tm = TEAM_ABBR.get(club)
        exact = [r for r in by_name.get(nm_norm(name), []) if id(r) not in taken]
        if len(exact) == 1:
            claim(exact[0], pick)
            continue
        if exact:                        # same name twice: the club decides
            if not any(claim(r, pick) for r in exact if r["team"] == tm):
                claim(min(exact, key=lambda r: num(r["age"]) or 99), pick)
            continue
        alias = FANTRAX_ALIAS_REV.get(nm_norm(name))
        pool = by_name.get(alias, []) if alias else [
            r for rs in by_name.values() for r in rs if same_person(r["name"], name)]
        for r in pool:                   # a variant spelling must match club
            if r["team"] == tm and claim(r, pick):
                break


def norm(n):
    n = unicodedata.normalize("NFKD", str(n))
    n = "".join(c for c in n if not unicodedata.combining(c))
    return re.sub(r"[^a-z]", "", n.lower())


def num(s):
    try:
        return float(str(s).replace(",", ""))
    except (TypeError, ValueError):
        return None


FAILED = object()           # the request did not come back; NOT "no record"


def fetch(url, refresh=False, delay=0.25):
    """GET and cache some JSON. Returns FAILED - never None - when the request
    itself fails, so a network problem can never be read as an empty result.
    The NHL edge 403s urllib's default user-agent, hence UA."""
    CACHE.mkdir(parents=True, exist_ok=True)
    f = CACHE / (re.sub(r"[^A-Za-z0-9]+", "-", url)[-120:].strip("-") + ".json")
    if f.exists() and not refresh:
        try:
            return json.loads(f.read_text())
        except json.JSONDecodeError:
            pass
    try:
        with urllib.request.urlopen(
                urllib.request.Request(url, headers=UA), timeout=30) as r:
            data = json.load(r)
    except (urllib.error.URLError, OSError, json.JSONDecodeError) as e:
        print(f"  ! {url}: {e}")
        return FAILED
    f.write_text(json.dumps(data))
    time.sleep(delay)
    return data


def first_last(name):
    parts = str(name).split()
    return (norm(parts[0]) if parts else ""), (norm(parts[-1]) if parts else "")


def nhl_id(name, team, refresh, delay):
    """Fantrax name -> NHL player id, or None when the NHL has no such player.

    Searched by SURNAME, because the endpoint matches on first name only:
    querying "Christopher Tanev" returns a list of Christophers and not one
    Tanev. The surname narrows it, then the first name picks the man - allowing
    for the forms the two sources disagree on (Chris/Christopher Tanev,
    Mikey/Michael Eyssimont, Alexei/Alexey Toropchenko, John/Jack St. Ivany).

    Returns AMBIGUOUS rather than None when the surname exists but no first
    name fits, so a failed match is never read as "never played in the NHL"."""
    first, last = first_last(name)
    hits = fetch(SEARCH + urllib.parse.quote(name.split()[-1]), refresh, delay)
    if hits is FAILED:
        return FAILED
    hits = [h for h in (hits or []) if first_last(h.get("name"))[1] == last]
    if not hits:
        return None                     # genuinely no NHL record
    good = [h for h in hits if same_first(first_last(h["name"])[0], first)]
    if not good and len(hits) == 1 and hits[0].get("active"):
        good = hits                     # only one active man by that surname
    if not good:
        # The surname belongs to some NHL player, but not to this one - an
        # Alexander Zetterberg is not Henrik. Ask again the other way round,
        # since the endpoint also matches on first name; if he is still not
        # there under his own surname, he has never played in the NHL.
        alt = fetch(SEARCH + urllib.parse.quote(name.split()[0]), refresh, delay)
        if alt is FAILED:
            return FAILED
        good = [h for h in (alt or []) if first_last(h.get("name")) == (first, last)]
        if not good:
            return None
    for h in good:                      # same club as the export says
        if team and (h.get("teamAbbrev") or h.get("lastTeamAbbrev")) == team:
            return h["playerId"]
    # namesakes span eras - there are two Colin Millers, one of whom last
    # played before the league kept season ids. Take the one still playing.
    good.sort(key=lambda h: (bool(h.get("active")),
                             str(h.get("lastSeasonId") or "")), reverse=True)
    return good[0]["playerId"]


def career_gp(pid, refresh, delay):
    d = fetch(LANDING.format(pid), refresh, delay)
    if d is FAILED or not d:
        return FAILED
    return d.get("careerTotals", {}).get("regularSeason", {}).get("gamesPlayed", 0)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", default="data")
    ap.add_argument("--min-proj", type=float, default=10.0,
                    help="keep free agents projected at least this (default 10)")
    ap.add_argument("--min-age", type=float, default=18.0,
                    help="rule guide: \"you cannot draft any player that is "
                         "not at the NHL Draft eligible age\" - 18 (default 18)")
    ap.add_argument("--max-age", type=float, default=23.0,
                    help="also keep every free agent this age or younger, "
                         "projection or not - most prospect-draft targets have "
                         "no Fantrax projection at all (default 23)")
    ap.add_argument("--delay", type=float, default=0.25)
    ap.add_argument("--workers", type=int, default=6,
                    help="parallel lookups (default 6)")
    ap.add_argument("--refresh", action="store_true")
    ap.add_argument("--out", default=str(OUT))
    args = ap.parse_args()

    d = pathlib.Path(args.data)
    act = {r["ID"].strip("*"): r
           for r in csv.DictReader((d / "fantrax-actual-2526.csv").open())}
    prj = {r["ID"].strip("*"): r
           for r in csv.DictReader((d / "fantrax-proj-2627.csv").open())}

    cands = []
    for pid, r in act.items():
        if r["Status"] != "FA":                      # already owned
            continue
        proj = num(prj.get(pid, {}).get("FPts")) or 0
        age = num(r["Age"])
        if (proj < args.min_proj and not r["Rookie"]
                and not (age is not None and age <= args.max_age)):
            continue
        cands.append((pid, r, proj))
    print(f"{len(cands)} free agents to check against the NHL API")

    def check(item):
        """One candidate -> one board row. Network-bound, so these run in a
        small thread pool; the on-disk cache makes a rerun nearly free."""
        pid, r, proj = item
        fp, rate = num(r["FPts"]), num(r["FP/G"])
        gp = round(fp / rate) if fp and rate else 0
        age = num(r["Age"])
        pos = r["Position"]
        cap = GP_CAP.get(pos, 40)
        nid = nhl_id(r["Player"], r["Team"], args.refresh, args.delay)
        if nid is FAILED:
            cgp, note, elig = "", "LOOKUP FAILED - not checked", "?"
        elif nid is AMBIGUOUS:
            cgp, note, elig = "", "NAME NOT MATCHED - check by hand", "?"
        elif nid is None:
            cgp, note, elig = 0, "no NHL record - import or junior", "YES"
        else:
            cgp = career_gp(nid, args.refresh, args.delay)
            if cgp is FAILED:
                cgp, note, elig = "", "LOOKUP FAILED - not checked", "?"
            elif gp and not cgp:
                # the export has him playing this season, the NHL record has
                # him never playing at all: the name matched the wrong man
                cgp, note, elig = "", "MISMATCH - 25/26 games but no NHL record", "?"
            else:
                note, elig = "", ("YES" if cgp <= cap else "no")
        # rule guide s."Prospect Draft": the games cap is not the only gate -
        # "you cannot draft any player that is not at the NHL Draft eligible
        # age". Landon Dupont is a 2027 name and cannot be taken this year.
        if (elig == "YES" and age is not None and age < args.min_age):
            elig = "no"
            note = f"under NHL draft age ({age:.0f})"
        return dict(name=r["Player"], team=r["Team"], pos=pos,
                    age=age, career_gp=cgp, gp_2526=gp,
                    fpts_2526=fp or 0, fpg_2526=rate or 0, proj_2627=proj,
                    eligible=elig, rookie=r["Rookie"], draft_pick="",
                    nhl_id="" if nid in (None, FAILED, AMBIGUOUS) else nid,
                    note=note)

    rows = []
    with concurrent.futures.ThreadPoolExecutor(args.workers) as pool:
        for i, row in enumerate(pool.map(check, cands), 1):
            rows.append(row)
            if i % 100 == 0:
                print(f"  ...{i}/{len(cands)}", flush=True)

    assign_picks(rows)
    # the NHL drafted these players, so they are at draft-eligible age whatever
    # birthday Fantrax has on file - Yuri Ivanov and Max Laatikainen both read 17
    for r in rows:
        if r["draft_pick"] and r["eligible"] == "no" \
                and r["note"].startswith("under NHL draft age"):
            r["eligible"], r["note"] = "YES", ""

    rows.sort(key=lambda x: ({"YES": 0, "?": 1}.get(x["eligible"], 2),
                             -x["proj_2627"]))
    out = pathlib.Path(args.out)
    with out.open("w", newline="") as f:
        w = csv.DictWriter(f, FIELDS)
        w.writeheader()
        w.writerows(rows)
    ok = [x for x in rows if x["eligible"] == "YES"]
    bad = [x for x in rows if x["eligible"] == "?"]
    print(f"wrote {out}  ({len(rows)} checked, {len(ok)} prospect-eligible)")
    if bad:
        print(f"  ! {len(bad)} marked '?' - not counted eligible; see the note column")
        for x in bad[:15]:
            print(f"      {x['name']:24s} {x['team']:6s} {x['note']}")
    for x in ok[:25]:
        print(f"  {x['name']:24s} {x['team']:4s} {x['pos']:2s} "
              f"career {x['career_gp']:>3} GP  proj {x['proj_2627']:>6.1f}  "
              f"{x['note']}")


if __name__ == "__main__":
    main()
