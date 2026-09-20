#!/usr/bin/env python3
"""Flatten the Fantrax pulls in ./data into a multi-sheet .xlsx.

The workbook is split in two.

SOURCE sheets are owned by this script and rewritten from scratch every run:
Player Pool is the master - one row per player, every stat, computed in python.
Rosters layers ownership over it (team, slot, roster status, cap hit) and pulls
each player's numbers back out of Player Pool by Fantrax ID, so a stat is
defined in exactly one place. It also carries the helper keys the rest of the
workbook navigates by.

DERIVED sheets - My Roster, Team Summary - are written once, entirely
out of formulas that look into the source sheets, and are then left alone. A
data refresh flows through them without disturbing anything typed in by hand
(the scenario blocks on My Roster, above all). Pass --rebuild-derived to
regenerate them; the names typed into the scenario blocks are carried across
even then.

Usage: ./build_workbook.py [--data data] [--out league.xlsx]
"""
import argparse
import csv
import json
import pathlib

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as gl
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

try:                                # hand-gathered, not a Fantrax pull
    from nhl_draft import DRAFT_YEAR, PICKS
except ImportError:
    DRAFT_YEAR, PICKS = "", []

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(color="FFFFFF", bold=True)
KEY_FILL = PatternFill("solid", fgColor="8497B0")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
SCEN_FILL = PatternFill("solid", fgColor="2E5C8A")
BAND_FILL = PatternFill("solid", fgColor="4A6FA5")
MONEY = '#,##0'
DEC2 = '0.00'
DEC3 = '0.000'
TOP_BORDER = Border(top=Side(style="thin"))

PRIOR = "25/26"   # the actuals export
PROJ = "26/27"    # the projections export

# Fixed section sizes on the derived sheets. The rule guide caps every one of
# them, so a section is laid out once and left to fill itself: 20 active
# (12F/6D/2G), 3 bench, 2 IR, 26 prospects. The prospect and drop allowances run
# past those caps because the offseason permits going over (s."Trades Reopen").
N_ACTIVE, N_BENCH, N_IR, N_MINORS = 20, 3, 2, 32
N_TEAMS = 20
N_SCEN = 3
SCENARIOS = True      # side-by-side what-if copies of the roster on My Roster

POOL = "Player Pool"
DRAFT_SHEET = f"NHL Draft {DRAFT_YEAR}"
BOARD_SHEET = "Prospect Board"
SOURCE_SHEETS = [POOL, "FA Pool", "Rosters", f"Projections {PROJ.replace('/', '-')}",
                 f"Last Season {PRIOR.replace('/', '-')}", "External Projections",
                 DRAFT_SHEET, BOARD_SHEET, "Notes", "Scoring",
                 "League Settings"]
DERIVED_SHEETS = ["My Roster", "Team Summary"]

# Player Pool: the master table. Column letters are fixed so Rosters can quote
# them; every entry here is a value computed in python.
P = {"player": "A", "nhl": "B", "pos": "C", "elig": "D", "age": "E",
     "nhlstatus": "F", "owner": "G", "salary": "H",
     "fpts": "I", "fptsperm": "J", "fpg": "K", "fpgperm": "L",
     "proj": "M", "projfpg": "N", "ag": "O", "cons": "P", "change": "Q",
     "consperm": "R", "pfpg": "S", "pfpgperm": "T", "adp": "U",
     "id": "V", "rw": "W", "si": "X", "sr": "Y", "label": "Z"}
P_HEADERS = [
    ("player", "Player", 24), ("nhl", "NHL", 6), ("pos", "Pos", 7),
    ("elig", "Eligible", 9), ("age", "Age", 6), ("nhlstatus", "NHL Status", 11),
    ("owner", "Owner", 24), ("salary", "Salary", 13),
    ("fpts", f"FPts {PRIOR}", 11), ("fptsperm", f"FPts {PRIOR} per $M", 13),
    ("fpg", f"FP/G {PRIOR}", 11), ("fpgperm", f"FP/G {PRIOR} per $M", 13),
    ("proj", f"Fantrax {PROJ}", 11), ("projfpg", "Fantrax FP/G", 12),
    ("ag", f"A&G {PROJ}", 11), ("cons", f"Consensus {PROJ}", 13),
    ("change", "Change", 10), ("consperm", "Cons. per $M", 12),
    ("pfpg", "Proj FP/G", 10), ("pfpgperm", "Proj FP/G per $M", 13),
    ("adp", "ADP", 8), ("id", "Fantrax ID", 11), ("rw", "Rotowire ID", 12),
    ("si", "StatsInc ID", 12), ("sr", "SportRadar ID", 38),
    ("label", "Label", 30)]

# Rosters: ownership over the master. A handful of columns are written by
# python because they ARE the ownership data; the rest look back into the pool.
R = {"team": "A", "player": "B", "nhl": "C", "pos": "D", "elig": "E", "slot": "F",
     "status": "G", "age": "H", "salary": "I", "fpts": "J", "fpg": "K",
     "proj": "L", "projfpg": "M", "change": "N", "consperm": "O", "adp": "P",
     "ag": "Q", "cons": "R", "fptsperm": "S", "fpgperm": "T", "pfpg": "U",
     "pfpgperm": "V", "id": "W",
     "poolrow": "X", "sortcons": "Y", "sortfpts": "Z", "teamno": "AA",
     "slotrank": "AB", "droprank": "AC", "slotkey": "AD", "dropkey": "AE",
     "teamkey": "AF"}
R_HEADERS = [
    ("team", "Fantasy Team", 24), ("player", "Player", 24), ("nhl", "NHL", 6),
    ("pos", "Pos", 7), ("elig", "Eligible", 9), ("slot", "Slot", 7),
    ("status", "Status", 16), ("age", "Age", 6), ("salary", "Salary", 13),
    ("fpts", f"FPts {PRIOR}", 11), ("fpg", f"FP/G {PRIOR}", 11),
    ("proj", f"Fantrax {PROJ}", 11), ("projfpg", "Fantrax FP/G", 12),
    ("change", "Change", 10), ("consperm", "Cons. per $M", 12), ("adp", "ADP", 8),
    ("ag", f"A&G {PROJ}", 11), ("cons", f"Consensus {PROJ}", 13),
    ("fptsperm", f"FPts {PRIOR} per $M", 13), ("fpgperm", f"FP/G {PRIOR} per $M", 13),
    ("pfpg", "Proj FP/G", 10), ("pfpgperm", "Proj FP/G per $M", 13),
    ("id", "Fantrax ID", 11),
    ("poolrow", "· pool row", 10), ("sortcons", "· sort cons", 11),
    ("sortfpts", "· sort fpts", 11), ("teamno", "· team #", 9),
    ("slotrank", "· slot rank", 11), ("droprank", "· drop rank", 11),
    ("slotkey", "· slot key", 26), ("dropkey", "· drop key", 22),
    ("teamkey", "· team key", 11)]
R_KEYS = 9                                  # trailing helper columns
R_PY = ("team", "slot", "status", "salary", "id")        # written, not looked up
R_LOOKUP = ["player", "nhl", "pos", "elig", "age", "fpts", "fpg", "proj",
            "projfpg", "change", "adp", "ag", "cons", "pfpg"]
R_PER_M = [("fptsperm", "fpts", DEC2), ("fpgperm", "fpg", DEC3),
           ("pfpgperm", "pfpg", DEC3), ("consperm", "cons", DEC2)]


def load(data_dir, name):
    return json.loads((data_dir / f"{name}.json").read_text())


def num(s):
    """Fantrax CSV numerics: '12,500,000' -> 12500000.0, '-'/'' -> None."""
    if s is None:
        return None
    s = s.strip().replace(",", "").replace("%", "")
    if s in ("", "-", "N/A"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _wrap_lines(text, width):
    """Lines a greedy word-wrap needs to fit `text` into `width` characters."""
    lines, cur = 1, 0
    for word in text.split():
        if cur == 0:
            cur = len(word)
        elif cur + 1 + len(word) <= width:
            cur += 1 + len(word)
        else:
            lines, cur = lines + 1, len(word)
    return lines


def header_width(text, max_lines, pad):
    """Narrowest column width that wraps the header into `max_lines` or fewer."""
    text = str(text)
    longest = max((len(w) for w in text.split()), default=len(text))
    for w in range(longest, len(text) + 1):
        if _wrap_lines(text, w) <= max_lines:
            return w + pad
    return len(text) + pad


def fit_headers(ws, headers, max_lines=2, pad=3):
    """Widen any column too narrow for its header, then pin the header height.

    The height matters as much as the width: headers are wrap_text, and both
    Excel and LibreOffice stamp a fixed 15pt row height when they save, which
    silently clips every two-line header on the next open. `pad` covers the
    cell inset plus the autofilter arrow."""
    lines = 1
    for i, h in enumerate(headers, start=1):
        need = header_width(h, max_lines, pad)
        width = ws.column_dimensions[gl(i)].width or 0
        if width < need:
            ws.column_dimensions[gl(i)].width = width = need
        lines = max(lines, _wrap_lines(str(h), max(int(width) - pad, 1)))
    ws.row_dimensions[1].height = 13.5 * lines + 5


def head_row(ws, headers, widths=None, row=1):
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.fill, c.font = HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(vertical="center", wrap_text=True)
        if widths and i - 1 < len(widths) and widths[i - 1]:
            ws.column_dimensions[gl(i)].width = widths[i - 1]


def add_sheet(wb, title, headers, rows, fmts=None, widths=None):
    """Plain data sheet: header, rows, freeze, autofilter, widths."""
    ws = wb.create_sheet(title)
    ws.append(headers)
    for c in ws[1]:
        c.fill, c.font = HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(vertical="center", wrap_text=True)
    for r in rows:
        ws.append(r)
    ws.freeze_panes = "A2"
    if rows:
        ws.auto_filter.ref = f"A1:{gl(len(headers))}{len(rows) + 1}"
    for idx, fmt in (fmts or {}).items():
        for row in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):
            row[0].number_format = fmt
    for i, h in enumerate(headers, start=1):
        if widths and i - 1 < len(widths) and widths[i - 1]:
            w = widths[i - 1]
        else:
            lens = [len(str(r[i - 1])) for r in rows[:500]
                    if i - 1 < len(r) and r[i - 1] is not None]
            w = min(max(max(lens + [len(str(h))]) + 2, 9), 42)
        ws.column_dimensions[gl(i)].width = w
    fit_headers(ws, headers)
    return ws


def blank_safe(idx):
    """INDEX returns 0 for an empty source cell; keep those blank instead."""
    return f'=IFERROR(IF({idx}="","",{idx}),"")'


def lookup(sheet, col, last, ptr):
    """One cell of `sheet`!col, addressed by a row-number cell."""
    q = f"'{sheet}'!" if " " in sheet else f"{sheet}!"
    return blank_safe(f"INDEX({q}${col}$2:${col}${last},{ptr})")


# ---------------------------------------------------------------- source sheets
def build_player_pool(wb, ctx):
    """The master table: one row per player in the pool, every stat computed."""
    ex, pids, pinfo, owner_of = ctx["ex"], ctx["pids"], ctx["pinfo"], ctx["owner_of"]
    fpts, proj, ag, cons = ctx["fpts"], ctx["proj"], ctx["ag"], ctx["cons"]
    per_m, rate_per_m, delta, fpg = (ctx["per_m"], ctx["rate_per_m"],
                                     ctx["delta"], ctx["fpg"])
    rows = []
    for pid, meta in pinfo.items():
        p, e = pids.get(pid, {}), ex.get(pid, {})
        sal = num(e.get("Salary"))
        fp, rate, pj, cn = fpts(pid), num(e.get("FP/G")), proj(pid), cons(pid)
        rows.append([
            ctx["name_of"](pid), e.get("Team") or p.get("team"), p.get("position"),
            meta.get("eligiblePos"), num(e.get("Age")), meta.get("status"),
            owner_of.get(pid, "") or "FA", sal,
            fp, per_m(fp, sal), rate, rate_per_m(rate, sal),
            pj, num(ctx["pr"].get(pid, {}).get("FP/G")), ag(pid), cn, delta(pid),
            per_m(cn, sal), fpg(pid), rate_per_m(fpg(pid), sal), num(e.get("ADP")),
            pid, p.get("rotowireId"), p.get("statsIncId"), p.get("sportRadarId"),
            None])
    rows.sort(key=lambda r: (-(r[15] or 0), -(r[8] or 0), r[0]))

    # A label unique enough to look a player up by. Plain names are not: the
    # pool carries 32 collisions, six of which share name, team AND position.
    seen, labels = {}, {}
    for r in rows:
        base = f"{r[0]} ({r[1] or '?'} {r[2] or '?'})"
        seen[base] = seen.get(base, 0) + 1
        r[-1] = base if seen[base] == 1 else f"{base} #{seen[base]}"
        labels[r[21]] = r[-1]
    ctx["labels"] = labels
    fmts = {}
    for key, fmt in (("salary", MONEY), ("fpts", DEC2), ("fptsperm", DEC2),
                     ("fpg", DEC2), ("fpgperm", DEC3), ("proj", DEC2),
                     ("projfpg", DEC2), ("ag", DEC2), ("cons", DEC2),
                     ("change", DEC2), ("consperm", DEC2), ("pfpg", DEC3),
                     ("pfpgperm", DEC3)):
        fmts[[k for k, _, _ in P_HEADERS].index(key) + 1] = fmt
    add_sheet(wb, POOL, [h for _, h, _ in P_HEADERS], rows, fmts=fmts,
              widths=[w for _, _, w in P_HEADERS])
    ctx["pool_last"] = len(rows) + 1


def build_fa_pool(wb, ctx):
    """Free agents carrying a projection: the FA-draft board, and the source
    for the pick-list dropdown on the scenario blocks.

    Anyone with no projection and no points last season is dropped - that is
    ~7,400 of the 8,100 unowned players, and none of them is a draft target.
    Column P holds the same labels sorted A-Z, sitting outside the autofilter
    so re-sorting the board cannot drag it along; the FAPool defined name
    points there and the dropdowns read it."""
    ex, pinfo, pids = ctx["ex"], ctx["pinfo"], ctx["pids"]
    fpts, proj, ag, cons = ctx["fpts"], ctx["proj"], ctx["ag"], ctx["cons"]
    rows = []
    for pid, meta in pinfo.items():
        if ctx["owner_of"].get(pid):
            continue
        pj, cn, fp = proj(pid), cons(pid), fpts(pid)
        if not (pj or cn or fp):
            continue
        e = ex.get(pid, {})
        sal = num(e.get("Salary"))
        rows.append([ctx["labels"].get(pid), ctx["name_of"](pid),
                     e.get("Team") or pids.get(pid, {}).get("team"),
                     meta.get("eligiblePos"), num(e.get("Age")), sal,
                     fp, pj, ag(pid), cn, ctx["per_m"](cn, sal),
                     num(e.get("ADP")), pid])
    rows.sort(key=lambda r: (-(r[9] or 0), -(r[7] or 0), r[1]))
    for i, r in enumerate(rows, start=1):
        r.insert(0, i)
    ws = add_sheet(wb, "FA Pool",
                   ["Rank", "Label", "Player", "NHL", "Pos", "Age", "Salary",
                    f"FPts {PRIOR}", f"Fantrax {PROJ}", f"A&G {PROJ}",
                    f"Consensus {PROJ}", "Cons. per $M", "ADP", "Fantrax ID"],
                   rows, fmts={7: MONEY, 8: DEC2, 9: DEC2, 10: DEC2, 11: DEC2,
                               12: DEC2},
                   widths=[7, 30, 24, 6, 7, 6, 13, 11, 11, 11, 13, 12, 8, 11])
    last = len(rows) + 1
    ws.auto_filter.ref = f"A1:N{last}"
    ws.cell(row=1, column=16, value="Pick list (A-Z)").fill = HEADER_FILL
    ws.cell(row=1, column=16).font = HEADER_FONT
    ws.column_dimensions["P"].width = 30
    for i, label in enumerate(sorted(r[1] for r in rows), start=2):
        ws.cell(row=i, column=16, value=label)

    name = "FAPool"
    if name in wb.defined_names:
        del wb.defined_names[name]
    wb.defined_names.add(DefinedName(name, attr_text=f"'FA Pool'!$P$2:$P${last}"))
    ctx["fa_last"] = last


def build_rosters(wb, ctx):
    """Ownership over the master, plus the keys the derived sheets navigate by.

    Only team / slot / roster status / cap hit / player ID are written here -
    they are the ownership facts. Everything else is an INDEX back into Player
    Pool, so no stat is stored twice and a pool refresh moves this sheet too.
    Salary stays local because it is the team's cap hit: cap retention (rule
    guide s."Cap Retention") makes it a property of the roster spot, not the
    player.

    The trailing keys give each row a rank within its team's roster section and
    a rank within the team by last season's points - the order the four
    mandatory drops are banded off - then string keys built from both. A derived
    row just MATCHes "<team>|ACTIVE|3" and INDEXes across."""
    rows, pool_last = ctx["roster_rows"], ctx["pool_last"]
    ws = wb.create_sheet("Rosters")
    heads = [h for _, h, _ in R_HEADERS]
    head_row(ws, heads, [w for _, _, w in R_HEADERS])
    for i in range(len(R_HEADERS) - R_KEYS + 1, len(R_HEADERS) + 1):
        ws.cell(row=1, column=i).fill = KEY_FILL

    idx = {k: i + 1 for i, (k, _, _) in enumerate(R_HEADERS)}
    last = len(rows) + 1
    for n, r in enumerate(rows, start=2):
        for key in R_PY:
            ws.cell(row=n, column=idx[key], value=r[key])
        ws.cell(row=n, column=idx["poolrow"],
                value=f'=IFERROR(MATCH(${R["id"]}{n},'
                      f"'{POOL}'!${P['id']}$2:${P['id']}${pool_last},0),\"\")")
        ptr = f'${R["poolrow"]}{n}'
        for key in R_LOOKUP:
            ws.cell(row=n, column=idx[key],
                    value=lookup(POOL, P[key], pool_last, ptr))
        for key, src, _ in R_PER_M:
            s, v = f'${R["salary"]}{n}', f'${R[src]}{n}'
            ws.cell(row=n, column=idx[key], value=(
                f'=IF(AND(ISNUMBER({v}),{v}<>0,ISNUMBER({s}),{s}>0),'
                f'{v}/({s}/1000000),"")'))

        A, G = R["team"], R["status"]
        X, Y = R["sortcons"], R["sortfpts"]
        ws.cell(row=n, column=idx["sortcons"],
                value=f'=IF(ISNUMBER(${R["cons"]}{n}),${R["cons"]}{n},-1)')
        ws.cell(row=n, column=idx["sortfpts"],
                value=f'=IF(ISNUMBER(${R["fpts"]}{n}),${R["fpts"]}{n},-1)')
        # Rosters is written team-major, so "same team as the row above" is all
        # the team counter needs - far cheaper than a distinct-count per row.
        ws.cell(row=n, column=idx["teamno"], value="=1" if n == 2 else
                f'=IF(${A}{n}=${A}{n-1},${R["teamno"]}{n-1},${R["teamno"]}{n-1}+1)')
        # rank within group; ties broken by sheet order, so the key stays unique
        ws.cell(row=n, column=idx["slotrank"], value=(
            f'=COUNTIFS(${A}$2:${A}${last},${A}{n},${G}$2:${G}${last},${G}{n},'
            f'${X}$2:${X}${last},">"&${X}{n})'
            f'+COUNTIFS(${A}$2:${A}{n},${A}{n},${G}$2:${G}{n},${G}{n},'
            f'${X}$2:${X}{n},${X}{n})'))
        ws.cell(row=n, column=idx["droprank"], value=(
            f'=IF(${G}{n}="MINORS","",'
            f'COUNTIFS(${A}$2:${A}${last},${A}{n},${G}$2:${G}${last},"<>MINORS",'
            f'${Y}$2:${Y}${last},">"&${Y}{n})'
            f'+COUNTIFS(${A}$2:${A}{n},${A}{n},${G}$2:${G}{n},"<>MINORS",'
            f'${Y}$2:${Y}{n},${Y}{n}))'))
        ws.cell(row=n, column=idx["slotkey"],
                value=f'=${A}{n}&"|"&${G}{n}&"|"&${R["slotrank"]}{n}')
        ws.cell(row=n, column=idx["dropkey"],
                value=f'=IF(${R["droprank"]}{n}="","",${A}{n}&"|"&${R["droprank"]}{n})')
        ws.cell(row=n, column=idx["teamkey"],
                value=f'=IF(COUNTIFS(${A}$2:${A}{n},${A}{n})=1,"T"&${R["teamno"]}{n},"")')

    for key, fmt in (("salary", MONEY), ("fpts", DEC2), ("fpg", DEC2),
                     ("proj", DEC2), ("projfpg", DEC2), ("change", DEC2),
                     ("consperm", DEC2), ("ag", DEC2), ("cons", DEC2),
                     ("fptsperm", DEC2), ("fpgperm", DEC3), ("pfpg", DEC3),
                     ("pfpgperm", DEC3)):
        for row in ws[f'{R[key]}2:{R[key]}{last}']:
            row[0].number_format = fmt
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{R['id']}{last}"
    fit_headers(ws, heads)
    ctx["roster_last"] = last


def season_sheet(wb, ctx, title, primary, secondary, plabel, slabel):
    """One ranked sheet per export, each carrying the other's number and the gap."""
    ex, pinfo, owner_of = ctx["ex"], ctx["pinfo"], ctx["owner_of"]
    rows = []
    for pid, e in ex.items():
        pv = primary(pid)
        if not pv:
            continue
        sal, sv = num(e["Salary"]), secondary(pid)
        rows.append([
            ctx["name_of"](pid), e["Team"], e["Position"],
            pinfo.get(pid, {}).get("eligiblePos"), num(e["Age"]),
            "Yes" if e["Rookie"].strip() else "", owner_of.get(pid, "") or "FA", sal,
            pv, sv, round(pv - sv, 2) if sv is not None else None,
            ctx["per_m"](pv, sal), num(e["ADP"]), num(e["Ros"]), pid])
    rows.sort(key=lambda r: -r[8])
    for i, r in enumerate(rows, start=1):
        r.insert(0, i)
    add_sheet(wb, title,
              ["Rank", "Player", "NHL", "Pos", "Eligible", "Age", "Rookie", "Owner",
               "Salary", plabel, slabel, "Change", f"{plabel} per $M",
               "ADP", "Rostered %", "Fantrax ID"],
              rows, fmts={9: MONEY, 10: DEC2, 11: DEC2, 12: DEC2, 13: DEC2},
              widths=[7, 24, 6, 7, 9, 6, 8, 24, 13, 12, 12, 10, 13, 8, 11, 11])


def build_external(wb, ctx):
    ext = ctx["ext"]
    if not ext:
        return
    srcs = ["A&G Avg", "A&G Nate", "A&G Blake"]
    ex, pinfo, owner_of = ctx["ex"], ctx["pinfo"], ctx["owner_of"]
    rows = []
    for pid, vals in ext.items():
        e = ex.get(pid, {})
        rows.append([ctx["name_of"](pid), e.get("Team"),
                     pinfo.get(pid, {}).get("eligiblePos"),
                     owner_of.get(pid, "") or "FA", num(e.get("Salary")),
                     ctx["fpts"](pid), ctx["proj"](pid)]
                    + [vals.get(k) for k in srcs]
                    + [ctx["cons"](pid),
                       round(vals["A&G Avg"] - ctx["proj"](pid), 2)
                       if vals.get("A&G Avg") and ctx["proj"](pid) else None, pid])
    rows.sort(key=lambda r: -(r[10] or 0))
    for i, r in enumerate(rows, start=1):
        r.insert(0, i)
    add_sheet(wb, "External Projections",
              ["Rank", "Player", "NHL", "Pos", "Owner", "Salary",
               f"Actual {PRIOR}", f"Fantrax {PROJ}"] + srcs
              + ["Consensus", "A&G - Fantrax", "Fantrax ID"],
              rows, fmts={6: MONEY, 7: DEC2, 8: DEC2, 9: DEC2, 10: DEC2,
                          11: DEC2, 12: DEC2, 13: DEC2},
              widths=[7, 24, 6, 7, 24, 13, 11, 12, 11, 11, 11, 12, 13, 11])


# ------------------------------------------------------------ NHL entry draft
TICK, UNTICK = "\u2611", "\u2610"          # checked / unchecked box glyphs
D_HEADS = ["\u2713", "Pick", "Rd", "Rd Pick", "Player", "Pos",
           "Dobber Upside", "NHL Certainty", "Dobber Fantasy Summary"]
D_WIDTHS = [5, 7, 5, 8, 26, 6, 13, 13, 92]
DEC1 = "0.0"


def read_ticks(wb, sheet):
    """Ticks typed into a board sheet, keyed by whatever column B holds -
    the draft pick number on one sheet, the player name on the other.

    Both are source sheets and so are rewritten from scratch every run; this
    is what carries the one hand-entered column across that rewrite."""
    if sheet not in wb.sheetnames:
        return set()
    return {row[1] for row in wb[sheet].iter_rows(min_row=2, max_col=2,
                                                  values_only=True)
            if row[1] is not None and row[0] == TICK}


def build_nhl_draft(wb, ctx):
    """Every pick of the last NHL entry draft, in the order they came off the
    board. Static research (nhl_draft.py), not a Fantrax pull - but rewritten
    each run like the other source sheets so the module stays the one copy.

    Column A is a click-to-tick box: a two-value dropdown rather than a real
    form control, which openpyxl cannot write and LibreOffice would not render.
    Filter on it to pull up the picks you have marked.

    The three Dobber columns come from import_dobber.py. Roughly one pick in
    seven - all of them late - has no DobberProspects page, and those are left
    blank rather than guessed at."""
    ticks, dob = ctx.get("draft_ticks", set()), ctx.get("dobber", {})
    rows = []
    for t in PICKS:
        d = dob.get(t[0], {})
        rows.append([TICK if t[0] in ticks else UNTICK] + list(t[:5])
                    + [num(d.get("upside")), num(d.get("certainty")),
                       d.get("summary") or None])
    ws = add_sheet(wb, DRAFT_SHEET, D_HEADS, rows, widths=D_WIDTHS,
                   fmts={7: DEC1, 8: DEC1})

    dv = DataValidation(type="list", formula1=f'"{UNTICK},{TICK}"',
                        allow_blank=True)
    dv.showErrorMessage = False
    ws.add_data_validation(dv)
    dv.add(f"A2:A{len(rows) + 1}")
    for r in range(2, len(rows) + 2):
        ws.cell(row=r, column=1).fill = INPUT_FILL
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=9).alignment = Alignment(wrap_text=True,
                                                       vertical="top")
    return ws


# -------------------------------------------------------------- prospect board
B_HEADS = ["\u2713", "Player", "Team", "Pos", "Age", "Career GP", "Elig",
           f"{DRAFT_YEAR} Pick", f"GP {PRIOR}", f"FPts {PRIOR}",
           f"FP/G {PRIOR}", f"Proj {PROJ}", "Dobber Upside", "NHL Certainty",
           "Dobber Fantasy Summary", "Note"]
B_WIDTHS = [5, 24, 6, 5, 6, 10, 6, 10, 8, 11, 10, 11, 13, 13, 92, 30]


def build_prospect_board(wb, ctx):
    """Free agents who still qualify as prospects, best projection first.

    Rule guide s."Prospect Draft (5 rounds)": a FWD/D who has yet to play his
    41st NHL game, a goalie his 26th. That is a CAREER count, which no Fantrax
    export carries, so prospect_board.py settles it against the NHL's own API
    and writes data/prospect-board.csv; this sheet is that file, joined to the
    Dobber ratings from import_dobber.py --players.

    Two gates, not one: the games cap, and s."Prospect Draft" again - "you
    cannot draft any player that is not at the NHL Draft eligible age", which
    is what keeps Landon Dupont (17, a 2027 name) off the board. Anyone the
    NHL actually drafted in 2026 is exempt from that age test by definition,
    and carries his pick number here so the draft class can be picked out.

    Anyone the NHL API has no record of has played no NHL games and is
    therefore eligible - the import bucket, arriving from the KHL, SHL, Liiga
    or the NCAA. The rule guide is explicit that he need not be a real-world
    rookie. Rows whose lookup failed are carried with Elig "?" rather than
    quietly counted as eligible."""
    ticks = ctx.get("board_ticks", set())

    def rank(r):
        """Prospect value first. These rows fill minor slots, which score no
        points, so what matters is what a player becomes - Dobber's upside -
        not what Fantrax expects him to put up next season. The projection is
        only the tie-breaker among players Dobber has not rated."""
        d = dobber_for(r)
        return (-(num(d.get("upside")) or 0), -(num(d.get("certainty")) or 0),
                -(num(r["proj_2627"]) or 0), r["name"])

    # The pool holds distinct players with identical names - an 18-year-old
    # Vladimir Dravecky drafted by the Islanders and a 41-year-old one long
    # retired. Ratings are keyed by name, so only one row may claim them:
    # the drafted player, else the youngest.
    same_name = {}
    for r in ctx.get("board", []):
        same_name.setdefault(r["name"], []).append(r)
    rated_row = {}
    for nm, rs in same_name.items():
        rated_row[nm] = id(min(rs, key=lambda r: (not r.get("draft_pick"),
                                                  num(r["age"]) or 99)))

    def dobber_for(r):
        if rated_row.get(r["name"]) != id(r):
            return {}
        return ctx.get("dobber_players", {}).get(r["name"], {})

    def worth_showing(r):
        """Eligible is not the same as worth a row. Every free agent under the
        games cap qualifies, which after the full sweep is 6,549 of them -
        including several hundred career minor-leaguers over 40 who are
        eligible on a technicality. A player earns his place here by having
        some signal: a Dobber rating, a projection, youth, or a draft pick.
        This hides rows, it does not judge eligibility - prospect-board.csv
        keeps every player and his verdict."""
        return bool(dobber_for(r).get("upside")) \
            or (num(r["proj_2627"]) or 0) > 0 \
            or (num(r["age"]) or 99) <= 23 or bool(r.get("draft_pick"))

    rows, hidden = [], 0
    for r in sorted(ctx.get("board", []), key=rank):
        if r.get("eligible") not in ("YES", "?"):
            continue
        if not worth_showing(r):
            hidden += 1
            continue
        d = dobber_for(r)
        rows.append([TICK if r["name"] in ticks else UNTICK,
                     r["name"], r["team"], r["pos"], num(r["age"]),
                     num(r["career_gp"]), r["eligible"],
                     num(r.get("draft_pick")), num(r["gp_2526"]),
                     num(r["fpts_2526"]), num(r["fpg_2526"]),
                     num(r["proj_2627"]), num(d.get("upside")),
                     num(d.get("certainty")), d.get("summary") or None,
                     r.get("note") or None])
    ws = add_sheet(wb, BOARD_SHEET, B_HEADS, rows, widths=B_WIDTHS,
                   fmts={10: DEC2, 11: DEC3, 12: DEC2, 13: DEC1, 14: DEC1})

    dv = DataValidation(type="list", formula1=f'"{UNTICK},{TICK}"',
                        allow_blank=True)
    dv.showErrorMessage = False
    ws.add_data_validation(dv)
    dv.add(f"A2:A{len(rows) + 1}")
    for r in range(2, len(rows) + 2):
        ws.cell(row=r, column=1).fill = INPUT_FILL
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=15).alignment = Alignment(wrap_text=True,
                                                        vertical="top")
    if hidden:
        c = ws.cell(row=len(rows) + 3, column=2, value=(
            f"{hidden} further players are prospect-eligible but carry no "
            f"rating, no projection and no draft pick - all are 24 or older. "
            f"They are in data/prospect-board.csv if you want them."))
        c.font = Font(italic=True, size=9)
    ctx["board_hidden"] = hidden
    return ws


def build_notes(wb, ctx):
    """Offseason research, looked up by player name from My Roster."""
    try:
        from drop_notes import NOTES
    except ImportError:
        NOTES = {}
    ws = add_sheet(wb, "Notes", ["Player", "Offseason news"],
                   sorted(NOTES.items()), widths=[24, 110])
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
        row[0].alignment = Alignment(wrap_text=True, vertical="top")
    ctx["notes_last"] = len(NOTES) + 1


def build_scoring(wb, ctx):
    scoring = []
    for group, cats in ctx["info"]["scoringSystem"]["scoringCategories"].items():
        for cat, bypos in cats.items():
            for pos, val in bypos.items():
                scoring.append([group.title(), cat, pos,
                                float(val.replace("points", ""))
                                if val.startswith("points") else val])
    scoring.sort(key=lambda r: (r[0], r[1], r[2] != "Default", r[2]))
    add_sheet(wb, "Scoring", ["Group", "Category", "Position", "Points"], scoring,
              widths=[12, 12, 11, 10])


def build_settings(wb, ctx):
    """Also the one place the derived sheets read 'which team is mine' from."""
    info, draft, ri = ctx["info"], ctx["draft"], ctx["info"]["rosterInfo"]
    settings = [
        ["My team", ctx["my_team"]],
        ["Salary cap", ctx["cap"]],
        ["League name", info["leagueName"]],
        ["Season", info["seasonYear"]],
        ["Season start", info["startDate"]],
        ["Season end", info["endDate"]],
        ["Scoring type", info["scoringSystem"]["type"]],
        ["Draft type", info["draftType"]],
        ["Draft date", draft.get("draftDate")],
        ["Draft state", draft.get("draftState")],
        ["Teams", len(ctx["teams"])],
        ["Roster spots", ri["maxTotalPlayers"]],
        ["Active spots", ri["maxTotalActivePlayers"]],
        ["Reserve spots", ri["maxTotalReservePlayers"]],
        ["Scoring periods", len(info.get("rosterPeriods", []))],
        ["Players in pool", len(ctx["pinfo"])],
        ["Players rostered", len(ctx["owner_of"])],
        [f"Players with {PRIOR} results",
         sum(1 for p in ctx["ex"] if num(ctx["ex"][p]["FPts"]))],
        [f"Players with {PROJ} projections",
         sum(1 for p in ctx["pr"] if num(ctx["pr"][p]["FPts"]))],
        ["Players with A&G projections",
         sum(1 for v in ctx["ext"].values() if v.get("A&G Avg"))],
        ["League ID (history)", info["leagueHistoryId"]],
    ]
    settings += [[f"Max active - {pos}", c["maxActive"]]
                 for pos, c in ri["positionConstraints"].items()]
    ws = add_sheet(wb, "League Settings", ["Setting", "Value"], settings,
                   widths=[26, 34])
    ws.cell(row=2, column=2).font = Font(bold=True)   # My team
    ws.cell(row=3, column=2).number_format = MONEY    # Salary cap


# --------------------------------------------------------------- derived sheets
TEAM_REF = "'League Settings'!$B$2"
CAP_REF = "'League Settings'!$B$3"

# (label, roster status, slots, note, counts against cap, scores)
MY_SECTIONS = [
    ("ROSTER", "ACTIVE", N_ACTIVE, "counts toward the cap and scores", True, True),
    ("RESERVE", "RESERVE", N_BENCH, "counts toward the cap, does NOT score", True, False),
    ("IR", "INJURED_RESERVE", N_IR, "outside the cap, does NOT score", False, False),
    ("MINORS", "MINORS", N_MINORS, "outside the cap, does NOT score", False, False),
]
MY_COLS = ["player", "nhl", "pos", "elig", "slot", "status", "age", "salary",
           "fpts", "fptsperm", "fpg", "fpgperm", "proj", "ag", "cons",
           "change", "consperm", "adp"]
MY_HEADS = ["Player", "NHL", "Pos", "Eligible", "Slot", "Status", "Age", "Salary",
            f"FPts {PRIOR}", f"FPts {PRIOR} per $M", f"FP/G {PRIOR}",
            f"FP/G {PRIOR} per $M", f"Fantrax {PROJ}", f"A&G {PROJ}",
            f"Consensus {PROJ}", "Change", "Cons. per $M", "ADP"]
MY_WIDTHS = [24, 6, 7, 9, 7, 16, 6, 13, 11, 13, 11, 13, 11, 11, 13, 10, 12, 8]
MY_FMT = {8: MONEY, 9: DEC2, 10: DEC2, 11: DEC2, 12: DEC3, 13: DEC2, 14: DEC2,
          15: DEC2, 16: DEC2, 17: DEC2}
SUM_COLS = (8, 9, 13, 14, 15)          # salary + the three points columns
PTR_COL = len(MY_COLS) + 1             # S: which Rosters row each slot resolves to

# One scenario block: the player name is typed, every other column resolves
# out of Player Pool by that name. Salary is the exception - see scen_salary.
SCEN_HEADS = [("Player", 26), ("Pos", 6), ("Age", 6), ("Salary", 13),
              (f"FPts {PRIOR}", 11), (f"Cons {PROJ}", 12), ("Cons/$M", 10),
              ("ADP", 8), ("Own", 8)]
SCEN_W = len(SCEN_HEADS) + 1           # +1 spacer column between blocks
SCEN_COL = PTR_COL + 1                 # T: first scenario block


def section_rows():
    """Where build_my_roster lays each section out.

    (label, status, note, first, last, subtotal, in_cap, scores). The scenario
    blocks are row-aligned with these, so a slot reads straight across from the
    live roster to all three what-ifs; read_scenarios uses the same rows to
    find the typed names again after a rebuild."""
    row, out = 1, []
    for label, status, slots, note, in_cap, scores in MY_SECTIONS:
        row += 2                            # blank line, then the band
        first = row + 1
        row += slots + 1                    # the slots, then the subtotal
        out.append((label, status, note, first, row - 1, row, in_cap, scores))
    return out


def _band(ws, row, c0, width, text):
    for c in range(c0, c0 + width):
        ws.cell(row=row, column=c).fill = BAND_FILL
    ws.cell(row=row, column=c0, value=text).font = Font(bold=True, color="FFFFFF")


def build_my_roster(wb, ctx):
    """My team, pulled out of Rosters a slot at a time.

    Each slot row resolves "<my team>|<status>|<n>" to a Rosters row, then every
    cell INDEXes across. Nothing here knows who is on the roster - change a team
    in Fantrax, re-pull, and the slots refill themselves."""
    last = ctx["roster_last"]
    ws = wb.create_sheet("My Roster")
    head_row(ws, MY_HEADS + ["· row"], MY_WIDTHS + [7])
    ws.cell(row=1, column=PTR_COL).fill = KEY_FILL
    ws.column_dimensions[gl(PTR_COL)].hidden = True
    # rows 1-2 stay put: row 1 is the header, row 2 the scenario title strip
    ws.freeze_panes = "A3"
    fit_headers(ws, MY_HEADS, pad=2)

    sects, sect_rows = section_rows(), {}
    for label, status, note, first, dlast, row, in_cap, scores in sects:
        _band(ws, first - 1, 1, len(MY_COLS), f"{label}  -  {note}")
        for k, r_ in enumerate(range(first, dlast + 1), start=1):
            ws.cell(row=r_, column=PTR_COL, value=(
                f'=IFERROR(MATCH({TEAM_REF}&"|{status}|"&{k},'
                f'Rosters!${R["slotkey"]}$2:${R["slotkey"]}${last},0),"")'))
            ptr = f"${gl(PTR_COL)}{r_}"
            for i, key in enumerate(MY_COLS, start=1):
                c = ws.cell(row=r_, column=i, value=lookup("Rosters", R[key], last, ptr))
                if i in MY_FMT:
                    c.number_format = MY_FMT[i]
        sect_rows[label] = (first, dlast, row, in_cap, scores)
        ws.cell(row=row, column=1, value=(
            f'="{label} subtotal  ("&COUNTIF(A{first}:A{dlast},"?*")&")"'))
        if in_cap:
            c = ws.cell(row=row, column=8, value=f"=SUM(H{first}:H{dlast})")
            c.number_format = MONEY
        if scores:
            for sc in SUM_COLS[1:]:
                L = gl(sc)
                ws.cell(row=row, column=sc,
                        value=f"=SUM({L}{first}:{L}{dlast})").number_format = DEC2
        for c_ in ws[row][:len(MY_COLS)]:
            c_.font = Font(bold=True)
            c_.border = TOP_BORDER
    row = sects[-1][5]

    cap_secs = [v for v in sect_rows.values() if v[3]]
    cap_rows = [v[2] for v in cap_secs]
    # count the cap sections one at a time: a single span would swallow the
    # band and subtotal rows sitting between them
    n_cap = "+".join(f'COUNTIF(A{f}:A{l},"?*")' for f, l, *_ in cap_secs)
    row += 2
    ws.cell(row=row, column=1, value="SALARY CAP").font = Font(bold=True)
    ws.cell(row=row, column=8, value=f"={CAP_REF}").number_format = MONEY
    r_used = row + 1
    ws.cell(row=r_used, column=1,
            value=f'="   Roster + Reserve ("&{n_cap}&") - counts against cap"')
    ws.cell(row=r_used, column=8,
            value="=" + "+".join(f"H{r}" for r in cap_rows)).number_format = MONEY
    r_space = r_used + 1
    ws.cell(row=r_space, column=1, value=(
        f'=IF({CAP_REF}-H{r_used}<0,"   OVER CAP by","   Cap space")')
        ).font = Font(bold=True)
    c = ws.cell(row=r_space, column=8, value=f"=ABS({CAP_REF}-H{r_used})")
    c.number_format, c.font = MONEY, Font(bold=True)
    # Rule guide s."Trades Reopen": offseason overage is allowed, and must be
    # cleared by the first day of the season.
    ws.conditional_formatting.add(f"H{r_space}", CellIsRule(
        operator="greaterThan", formula=[f"{CAP_REF}-H{r_used}"],
        font=Font(bold=True, color="C00000")))
    ws.cell(row=r_space + 1, column=1,          # $B$6 is "Season start"
            value='="   Offseason - overage permitted until "'
                  "&'League Settings'!$B$6").font = Font(italic=True)
    r_pts = r_space + 3
    ws.cell(row=r_pts, column=1, value="PROJECTED TEAM POINTS (Roster only)"
            ).font = Font(bold=True)
    ws.cell(row=r_pts, column=8,
            value=f"=O{sect_rows['ROSTER'][2]}").number_format = DEC2
    ctx["my_pts_row"] = sect_rows["ROSTER"][2]   # scenarios compare against it
    if SCENARIOS:
        build_scenarios(ws, ctx)


def seed_scenarios(ctx):
    """My roster as plain labels, bucketed and ordered the way My Roster
    shows it.

    These are written into the blocks as text, not formulas, so every one of
    them can be deleted or typed over. That is the whole trick: a scenario row
    is a name and nothing else, and the name is what the stats hang off."""
    mine = [r for r in ctx["roster_rows"] if r["team"] == ctx["my_team"]]
    labels, cons = ctx["labels"], ctx["cons"]

    def rank(fn):
        return lambda r: (-(fn(r["id"]) if fn(r["id"]) is not None else -1),
                          labels.get(r["id"]) or "")

    by_status = {}
    for _, status, *_ in MY_SECTIONS:
        rows = sorted((r for r in mine if r["status"] == status), key=rank(cons))
        by_status[status] = [labels.get(r["id"]) for r in rows]
    return by_status


def read_scenarios(wb):
    """The names typed into the scenario blocks, keyed by (block, row).

    My Roster survives a normal refresh untouched, so this only matters under
    --rebuild-derived - but that is exactly the run where losing a draft plan
    would hurt. Only the slot rows are read; bands, subtotals and the summary
    labels share the column and are regenerated."""
    if "My Roster" not in wb.sheetnames:
        return {}
    ws = wb["My Roster"]
    slots = {r for *_, first, last, _, _, _ in section_rows()
             for r in range(first, last + 1)}
    out = {}
    for k in range(N_SCEN):
        c = SCEN_COL + k * SCEN_W
        for r in slots:
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip() and not v.startswith("="):
                out[(k, r)] = v
    return out


def build_scenarios(ws, ctx):
    """Three what-if copies of the roster, side by side with the live one.

    Each block is row-aligned with the sections on the left, so a slot reads
    straight across: what you have, then the three plans for it. Exactly one
    cell per row is typed - the player's name, off a dropdown - and Pos, Age,
    Salary, points, Cons/$M, ADP and who owns him all resolve out of Player
    Pool from it. Clear the name and the row empties and every total moves;
    type a new one into any blank row and it joins them.

    Blocks start seeded with the current roster, so a plan begins as "what I
    have" and is edited down. Salary prefers this league's cap hit off Rosters
    over the pool's number, because cap retention (rule guide s."Cap
    Retention") makes the hit a property of the roster spot."""
    plast, rlast = ctx["pool_last"], ctx["roster_last"]
    seed = seed_scenarios(ctx)
    saved = ctx.get("scen_saved") or {}
    # a block that has been typed in is the user's plan, and is restored whole:
    # re-seeding it would quietly undo every player they had deleted
    edited = {k for k, _ in saved}
    sects = section_rows()
    base = sects[-1][5] + 2                 # the live sheet's cap block

    for k in range(N_SCEN):
        c0 = SCEN_COL + k * SCEN_W
        c_name, c_pos, c_age, c_sal, c_fpts, c_cons, c_cpm, c_adp, c_own = \
            range(c0, c0 + len(SCEN_HEADS))
        NM, SL = gl(c_name), gl(c_sal)

        def rg(c, a, b):
            L = gl(c)
            return f"${L}${a}:${L}${b}"

        for i, (h, w) in enumerate(SCEN_HEADS):
            cell = ws.cell(row=2, column=c0 + i, value=h)
            cell.fill, cell.font = HEADER_FILL, HEADER_FONT
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            ws.column_dimensions[gl(c0 + i)].width = w
        ws.column_dimensions[gl(c0 + len(SCEN_HEADS))].width = 2
        for c in range(c0, c0 + len(SCEN_HEADS)):
            ws.cell(row=1, column=c).fill = SCEN_FILL
        ws.cell(row=1, column=c_name, value=f"SCENARIO {k + 1}").font = Font(
            bold=True, color="FFFFFF", size=12)
        ws.cell(row=1, column=c_sal, value="<name this plan>").font = Font(
            italic=True, color="FFFFFF")

        dv = DataValidation(type="list", formula1="FAPool", allow_blank=True)
        dv.showErrorMessage = False      # a name off the FA board is still fine
        dv.promptTitle = "Add a player"
        dv.prompt = ("Pick a free agent, or type any label from Player Pool. "
                     "Clear the cell to drop the player.")
        ws.add_data_validation(dv)

        for label, status, note, first, dlast, sub, in_cap, scores in sects:
            _band(ws, first - 1, c0, len(SCEN_HEADS), f"{label}  -  {note}")
            dv.add(f"{NM}{first}:{NM}{dlast}")
            names = seed.get(status, [])
            for i, r_ in enumerate(range(first, dlast + 1)):
                nm = f"${NM}{r_}"
                m = f"MATCH({nm},'{POOL}'!${P['label']}$2:${P['label']}${plast},0)"
                val = (saved.get((k, r_)) if k in edited
                       else (names[i] if i < len(names) else None))
                cell = ws.cell(row=r_, column=c_name, value=val)
                cell.fill = INPUT_FILL
                for col, key, fmt in ((c_pos, "elig", None), (c_age, "age", None),
                                      (c_fpts, "fpts", DEC2), (c_cons, "cons", DEC2),
                                      (c_cpm, "consperm", DEC2), (c_adp, "adp", DEC2)):
                    c = ws.cell(row=r_, column=col, value=lookup(POOL, P[key], plast, m))
                    if fmt:
                        c.number_format = fmt
                # the pool salary is the league number; the cap hit on Rosters
                # is what this team actually pays, and wins where it exists
                pid = f"INDEX('{POOL}'!${P['id']}$2:${P['id']}${plast},{m})"
                hit = (f'IFERROR(INDEX(Rosters!${R["salary"]}$2:${R["salary"]}${rlast},'
                       f'MATCH({pid},Rosters!${R["id"]}$2:${R["id"]}${rlast},0)),"")')
                pool_sal = f"INDEX('{POOL}'!${P['salary']}$2:${P['salary']}${plast},{m})"
                ws.cell(row=r_, column=c_sal, value=(
                    f'=IFERROR(IF({nm}="","",IF(ISNUMBER({hit}),{hit},{pool_sal})),"")'
                )).number_format = MONEY
                own = f"INDEX('{POOL}'!${P['owner']}$2:${P['owner']}${plast},{m})"
                ws.cell(row=r_, column=c_own, value=(
                    f'=IFERROR(IF({nm}="","",IF({own}="FA","FA",'
                    f'IF({own}={TEAM_REF},"MINE","TAKEN"))),"")')
                ).alignment = Alignment(horizontal="center")

            ws.cell(row=sub, column=c_name, value=(
                f'="{label}  ("&COUNTIF({NM}{first}:{NM}{dlast},"?*")&")"'))
            for col, fmt in ((c_sal, MONEY), (c_fpts, DEC2), (c_cons, DEC2)):
                L = gl(col)
                ws.cell(row=sub, column=col,
                        value=f"=SUM({L}{first}:{L}{dlast})").number_format = fmt
            for c in range(c0, c0 + len(SCEN_HEADS)):
                ws.cell(row=sub, column=c).font = Font(bold=True)
                ws.cell(row=sub, column=c).border = TOP_BORDER

        S = {lab: (f, l, s) for lab, _, _, f, l, s, _, _ in sects}
        top, bot = sects[0][3], sects[-1][4]
        span = f"{NM}{top}:{NM}{bot}"
        slots = f"${NM}${top}:${NM}${bot}"      # every slot row in this block
        # a name typed twice in one block is almost always a paste slip
        ws.conditional_formatting.add(span, FormulaRule(
            formula=[f'AND({NM}{top}<>"",COUNTIF({slots},{NM}{top})>1)'],
            font=Font(bold=True, color="C00000")))
        own_span = (f"{gl(c_own)}{sects[0][3]}:{gl(c_own)}{sects[-1][4]}")
        ws.conditional_formatting.add(own_span, CellIsRule(
            operator="equal", formula=['"TAKEN"'], font=Font(color="C00000")))
        ws.conditional_formatting.add(own_span, CellIsRule(
            operator="equal", formula=['"FA"'], font=Font(bold=True, color="1F7A1F")))

        rr = base

        def band(text):
            nonlocal rr
            _band(ws, rr, c0, len(SCEN_HEADS), text)
            rr += 1

        def line(label, value, fmt=None, bold=False, check=None):
            nonlocal rr
            ws.cell(row=rr, column=c_name, value=label).font = Font(bold=bold)
            c = ws.cell(row=rr, column=c_sal, value=value)
            c.font = Font(bold=bold)
            if fmt:
                c.number_format = fmt
            if check:
                ws.cell(row=rr, column=c_cons, value=check(f"{SL}{rr}")
                        ).alignment = Alignment(horizontal="center")
            rr += 1
            return rr - 1

        def need(n):
            return lambda v: (f'=IF({v}={n},"OK",IF({v}<{n},"open "&({n}-{v}),'
                              f'"OVER "&({v}-{n})))')

        def cap(n):
            return lambda v: f'=IF({v}<={n},"OK","OVER "&({v}-{n}))'

        band("CAP")
        used = "=" + "+".join(f"{SL}{S[lab][2]}" for lab, *_, in_cap, _ in sects
                              if in_cap)
        r_used = line("Salary used  (roster + reserve)", used, MONEY)
        line("Salary cap", f"={CAP_REF}", MONEY)
        r_sp = line("Cap space  (negative = over)", f"={CAP_REF}-{SL}{r_used}",
                    MONEY, bold=True)
        for op, colour in (("lessThan", "C00000"), ("greaterThanOrEqual", "1F7A1F")):
            ws.conditional_formatting.add(f"{SL}{r_sp}", CellIsRule(
                operator=op, formula=["0"], font=Font(bold=True, color=colour)))

        rr += 1
        band("ACTIVE ROSTER")
        a_first, a_last, _ = S["ROSTER"]
        # dual eligibility ("D,F") counts toward both, which is what it buys
        for p, lab, n in (("F", "Forwards", 12), ("D", "Defence", 6),
                          ("G", "Goalies", 2)):
            line(f"{lab}  (need {n})",
                 f'=COUNTIF({rg(c_pos, a_first, a_last)},"*{p}*")', check=need(n))
        line("Active total  (need 20)",
             f'=COUNTIF({rg(c_name, a_first, a_last)},"?*")', bold=True,
             check=need(N_ACTIVE))
        for lab, key, n, chk in (("Reserve", "RESERVE", N_BENCH, cap(N_BENCH)),
                                 ("IR", "IR", N_IR, cap(N_IR)),
                                 ("Prospects", "MINORS", 26, cap(26))):
            f_, l_, _ = S[key]
            line(f"{lab}  (max {n})",
                 f'=COUNTIF({rg(c_name, f_, l_)},"?*")', check=chk)

        rr += 1
        band("POINTS")
        r_p = line(f"Projected points, active  (consensus {PROJ})",
                   f"={gl(c_cons)}{S['ROSTER'][2]}", DEC2, bold=True)
        line("Change vs today", f"={SL}{r_p}-$O${ctx['my_pts_row']}", DEC2)
        line(f"FPts {PRIOR}, active", f"={gl(c_fpts)}{S['ROSTER'][2]}", DEC2)

        rr += 1
        for msg in ("Type or pick a name to add a player; clear the cell to drop him.",
                    "Everything else on the row fills itself from Player Pool.",
                    "Own: FA = unowned, TAKEN = on another roster, MINE = yours.",
                    "Move a player between sections by cutting the name to another row.",
                    "Salary uses this league's cap hit where there is one, else the "
                    "pool number."):
            ws.cell(row=rr, column=c_name, value=msg).font = Font(italic=True, size=9)
            rr += 1


def build_team_summary(wb, ctx):
    """Every team, aggregated straight off Rosters with SUMIFS/COUNTIFS."""
    last, statuses = ctx["roster_last"], ctx["statuses"]
    heads = (["Rank", "Fantasy Team", "Players"]
             + [s.title().replace("_", " ") for s in statuses]
             + ["Total Salary", "Avg Salary", "Top Salary", f"Roster FPts {PRIOR}",
                f"Played {PRIOR}", f"Roster Proj {PROJ}", f"Projected {PROJ}",
                "Change", "Proj per $M"])
    ws = wb.create_sheet("Team Summary")
    head_row(ws, heads, [7, 24, 9] + [11] * len(statuses)
             + [14, 13, 13, 13, 11, 13, 13, 11, 12])
    ws.freeze_panes = "C2"
    fit_headers(ws, heads, pad=2)

    base = 4 + len(statuses)            # first column after the status counts
    C = {k: gl(base + i) for i, k in enumerate(
        ["salary", "avg", "top", "fpts", "played", "proj", "projn", "change", "perm"])}

    def src(key, a, b):
        return f"Rosters!${R[key]}${a}:${R[key]}${b}"

    for i in range(1, N_TEAMS + 1):
        r = i + 1
        t = f"$B{r}"
        ws.cell(row=r, column=2, value=(
            f'=IFERROR(INDEX({src("team", 2, last)},'
            f'MATCH("T"&{i},{src("teamkey", 2, last)},0)),"")'))
        guard = f'IF({t}="","",'
        ws.cell(row=r, column=3, value=f'={guard}COUNTIFS({src("team", 2, last)},{t}))')
        for j, st in enumerate(statuses):
            ws.cell(row=r, column=4 + j, value=(
                f'={guard}COUNTIFS({src("team", 2, last)},{t},'
                f'{src("status", 2, last)},"{st}"))'))
        ws.cell(row=r, column=base, value=(
            f'={guard}SUMIFS({src("salary", 2, last)},{src("team", 2, last)},{t}))')
            ).number_format = MONEY
        ws.cell(row=r, column=base + 1, value=(
            f'={guard}IFERROR({C["salary"]}{r}/C{r},""))')).number_format = MONEY
        ws.cell(row=r, column=base + 2, value=(
            f'={guard}SUMPRODUCT(MAX(({src("team", 2, last)}={t})'
            f'*{src("salary", 2, last)})))')).number_format = MONEY
        ws.cell(row=r, column=base + 3, value=(
            f'={guard}SUMIFS({src("fpts", 2, last)},{src("team", 2, last)},{t}))')
            ).number_format = DEC2
        ws.cell(row=r, column=base + 4, value=(
            f'={guard}COUNTIFS({src("team", 2, last)},{t},'
            f'{src("fpts", 2, last)},">0"))'))
        ws.cell(row=r, column=base + 5, value=(
            f'={guard}SUMIFS({src("proj", 2, last)},{src("team", 2, last)},{t}))')
            ).number_format = DEC2
        ws.cell(row=r, column=base + 6, value=(
            f'={guard}COUNTIFS({src("team", 2, last)},{t},'
            f'{src("proj", 2, last)},">0"))'))
        ws.cell(row=r, column=base + 7, value=(
            f'={guard}{C["proj"]}{r}-{C["fpts"]}{r})')).number_format = DEC2
        ws.cell(row=r, column=base + 8, value=(
            f'={guard}IFERROR({C["proj"]}{r}/({C["salary"]}{r}/1000000),""))')
            ).number_format = DEC2
        ws.cell(row=r, column=1, value=(
            f'={guard}RANK({C["proj"]}{r},'
            f'{C["proj"]}$2:{C["proj"]}${N_TEAMS + 1}))'))
    ws.auto_filter.ref = f"A1:{gl(len(heads))}{N_TEAMS + 1}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", default="data")
    ap.add_argument("--actual", default="data/fantrax-actual-2526.csv",
                    help=f"Fantrax player export holding {PRIOR} results")
    ap.add_argument("--proj", default="data/fantrax-proj-2627.csv",
                    help=f"Fantrax player export holding {PROJ} projections")
    ap.add_argument("--external", default="data/proj-external.csv",
                    help="third-party projections from import_projections.py")
    ap.add_argument("--dobber", default="data/dobber.csv",
                    help="DobberProspects ratings from import_dobber.py")
    ap.add_argument("--board", default="data/prospect-board.csv",
                    help="prospect-eligible free agents from prospect_board.py")
    ap.add_argument("--dobber-players", default="data/dobber-prospects.csv",
                    help="DobberProspects ratings for the board, by name")
    ap.add_argument("--team", help="your fantasy team name (default: from getLeagues.json)")
    ap.add_argument("--out", default="league.xlsx")
    ap.add_argument("--rebuild-derived", action="store_true",
                    help="regenerate My Roster / Team Summary, discarding "
                         "anything typed into the scenario blocks")
    ap.add_argument("--fresh", action="store_true",
                    help="ignore any existing workbook and build from scratch")
    args = ap.parse_args()

    d = pathlib.Path(args.data)
    info = load(d, "getLeagueInfo")
    rosters = load(d, "getTeamRosters")
    draft = load(d, "getDraftResults")
    pids = load(d, "getPlayerIds")
    pinfo = info["playerInfo"]

    # Fantrax CSV exports, keyed by Fantrax player ID ("*02un4*" -> "02un4").
    # Both carry identical roster metadata; only the FPts/FP/G columns differ.
    def read_export(path):
        f = pathlib.Path(path)
        if not f.exists():
            return {}
        return {r["ID"].strip("*"): r for r in csv.DictReader(f.open())}

    ex, pr = read_export(args.actual), read_export(args.proj)

    # DobberProspects ratings for the entry-draft class, keyed by draft pick
    dobber = {}
    df = pathlib.Path(args.dobber)
    if df.exists():
        dobber = {int(r["pick"]): r for r in csv.DictReader(df.open())}

    # the prospect-eligible free agents, and their Dobber ratings by name
    board, dobber_players = [], {}
    bf = pathlib.Path(args.board)
    if bf.exists():
        board = list(csv.DictReader(bf.open()))
    pf = pathlib.Path(args.dobber_players)
    if pf.exists():
        dobber_players = {r["player"]: r for r in csv.DictReader(pf.open())}

    # third-party projections, already scored with this league's rules
    ext, ext_gp = {}, {}
    xf = pathlib.Path(args.external)
    if xf.exists():
        for r in csv.DictReader(xf.open()):
            ext.setdefault(r["fantraxId"], {})[r["source"]] = num(r["FPts"])
            ext_gp.setdefault(r["fantraxId"], {})[r["source"]] = num(r.get("GP"))
    if not ex:                      # fall back so either export alone still builds
        ex, pr = pr, {}

    def name_of(pid):
        if pid in ex:
            return ex[pid]["Player"]
        n = pids.get(pid, {}).get("name", pid)
        if ", " in n:                       # "Madden, Tyler" -> "Tyler Madden"
            lastn, _, first = n.partition(", ")
            return f"{first} {lastn}"
        return n

    def fpts(pid):
        return num(ex.get(pid, {}).get("FPts"))

    def proj(pid):
        return num(pr.get(pid, {}).get("FPts"))

    def ag(pid):
        return ext.get(pid, {}).get("A&G Avg")

    def fpg(pid):
        """Blended per-game rate. Separates scoring rate from games played,
        which total projections conflate for anyone carrying an injury."""
        vals = [v for v in (num(pr.get(pid, {}).get("FP/G")),
                            (ag(pid) / ext_gp.get(pid, {}).get("A&G Avg")
                             if ag(pid) and ext_gp.get(pid, {}).get("A&G Avg") else None))
                if v]
        return round(sum(vals) / len(vals), 3) if vals else None

    def cons(pid):
        """Mean of the sources that cover this player. They are NOT rescaled:
        A&G runs ~6% hot vs Fantrax (84-game season, and no SHP/GWG/hat-tricks
        in their categories), so treat the spread as real disagreement."""
        vals = [v for v in (proj(pid), ag(pid)) if v]
        return round(sum(vals) / len(vals), 2) if vals else None

    def per_m(pts, sal):
        return round(pts / (sal / 1e6), 2) if pts and sal else None

    def rate_per_m(rate, sal):
        """A rate normalised to cap hit. Unlike total points per $M this is blind
        to games played, so an injured star is not punished twice."""
        return round(rate / (sal / 1e6), 3) if rate and sal else None

    def delta(pid):
        a, b = fpts(pid), proj(pid)
        return round(b - a, 2) if a is not None and b is not None else None

    owner_of, roster_rows = {}, []
    for team in rosters["rosters"].values():
        for it in team["rosterItems"]:
            owner_of[it["id"]] = team["teamName"]
            roster_rows.append({"team": team["teamName"], "slot": it["position"],
                                "status": it["status"], "salary": it.get("salary"),
                                "id": it["id"]})
    # team-major, best projection first: the team counter on Rosters relies on
    # every team's rows being contiguous
    roster_rows.sort(key=lambda r: (r["team"], -(proj(r["id"]) or 0)))

    my_team = args.team
    if not my_team:
        lg = d / "getLeagues.json"
        if lg.exists():
            leagues = json.loads(lg.read_text()).get("leagues", [])
            match = [l for l in leagues if l.get("leagueName") == info["leagueName"]]
            if match or leagues:
                my_team = (match or leagues)[0].get("teamName")
    known = {r["team"] for r in roster_rows}
    if my_team and my_team not in known:
        print(f"  ! team {my_team!r} not found; known teams: {sorted(known)}")

    ctx = dict(info=info, rosters=rosters, draft=draft, pids=pids, pinfo=pinfo,
               ex=ex, pr=pr, ext=ext, ext_gp=ext_gp, dobber=dobber,
               board=board, dobber_players=dobber_players,
               owner_of=owner_of,
               roster_rows=roster_rows, my_team=my_team,
               teams={t["id"]: t["name"] for t in info["teamInfo"].values()},
               statuses=sorted({r["status"] for r in roster_rows}),
               cap=next((t.get("salaryCap") for t in rosters["rosters"].values()), None),
               name_of=name_of, fpts=fpts, proj=proj, ag=ag, fpg=fpg, cons=cons,
               per_m=per_m, rate_per_m=rate_per_m, delta=delta)

    out = pathlib.Path(args.out)
    reuse = out.exists() and not args.fresh
    wb = load_workbook(out) if reuse else Workbook()
    if not reuse:
        wb.remove(wb.active)

    ctx["draft_ticks"] = read_ticks(wb, DRAFT_SHEET)
    ctx["board_ticks"] = read_ticks(wb, BOARD_SHEET)
    ctx["scen_saved"] = read_scenarios(wb)
    for name in list(wb.sheetnames):
        if name in SOURCE_SHEETS or (args.rebuild_derived and name in DERIVED_SHEETS):
            del wb[name]

    build_player_pool(wb, ctx)
    build_fa_pool(wb, ctx)
    build_rosters(wb, ctx)
    season_sheet(wb, ctx, f"Projections {PROJ.replace('/', '-')}", proj, fpts,
                 f"Proj {PROJ}", f"Actual {PRIOR}")
    season_sheet(wb, ctx, f"Last Season {PRIOR.replace('/', '-')}", fpts, proj,
                 f"FPts {PRIOR}", f"Proj {PROJ}")
    build_external(wb, ctx)
    build_nhl_draft(wb, ctx)
    build_prospect_board(wb, ctx)
    build_notes(wb, ctx)
    build_scoring(wb, ctx)
    build_settings(wb, ctx)

    made = []
    for name, fn in (("My Roster", build_my_roster),
                     ("Team Summary", build_team_summary)):
        if name not in wb.sheetnames:
            fn(wb, ctx)
            made.append(name)

    order = [n for n in DERIVED_SHEETS + SOURCE_SHEETS if n in wb.sheetnames]
    wb._sheets = ([wb[n] for n in order]
                  + [s for s in wb._sheets if s.title not in order])
    wb.save(out)

    kept = [n for n in DERIVED_SHEETS if n in wb.sheetnames and n not in made]
    print(f"wrote {out}  ({PRIOR} rows: {len(ex)}, {PROJ} rows: {len(pr)})")
    print(f"  sources rebuilt: {', '.join(SOURCE_SHEETS)}")
    if made:
        print(f"  derived created: {', '.join(made)}")
    if kept:
        print(f"  derived kept as-is (formulas, your input preserved): "
              f"{', '.join(kept)}")


if __name__ == "__main__":
    main()
