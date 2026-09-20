#!/usr/bin/env python3
"""Pull league data from the Fantrax FXEA API into ./data/*.json

Usage:  ./pull_league.py [leagueId] [--secret USER_SECRET_ID]

Both values default to LEAGUE_ID and FANTRAX_SECRET, read from ./.env or the
environment. The league ID is the token in your league URL:
  https://www.fantrax.com/fantasy/league/<leagueId>/home
The secret ID (Fantrax profile -> "Your Secret ID") is only needed for private
leagues or user-scoped calls.
"""
import argparse
import json
import os
import pathlib
import sys
import urllib.parse
import urllib.request

BASE = "https://www.fantrax.com/fxea/general"

# endpoint -> (extra query params, scope)
#   scope "league" -> takes leagueId;  scope "user" -> takes userSecretId only
ENDPOINTS = {
    "getLeagueInfo": ({}, "league"),
    "getTeamRosters": ({}, "league"),
    "getStandings": ({}, "league"),
    "getDraftResults": ({}, "league"),
    "getPlayerIds": ({"sport": "NHL"}, "league"),
    "getLeagues": ({}, "user"),
}


def load_dotenv(path=".env"):
    """Minimal .env reader: KEY=value lines, # comments, optional quotes."""
    f = pathlib.Path(path)
    if not f.exists():
        return
    for line in f.read_text().splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, val = line.partition("=")
        os.environ.setdefault(key.strip(), val.strip().strip("\"'"))


def fetch(endpoint, params):
    url = f"{BASE}/{endpoint}?{urllib.parse.urlencode(params)}"
    req = urllib.request.Request(url, headers={"User-Agent": "keeper-league/0.1"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read())


def main():
    load_dotenv()

    ap = argparse.ArgumentParser()
    ap.add_argument("league_id", nargs="?", default=os.environ.get("LEAGUE_ID"))
    ap.add_argument("--secret", default=os.environ.get("FANTRAX_SECRET"),
                    help="user secret ID, for private leagues")
    ap.add_argument("--out", default="data", help="output directory")
    args = ap.parse_args()

    if not args.league_id:
        ap.error("no league ID given and LEAGUE_ID not set in .env")

    outdir = pathlib.Path(args.out)
    outdir.mkdir(exist_ok=True)

    failures = 0
    for endpoint, (extra, scope) in ENDPOINTS.items():
        if scope == "user":
            if not args.secret:
                continue
            params = {"userSecretId": args.secret, **extra}
        else:
            params = {"leagueId": args.league_id, **extra}
            if args.secret:
                params["userSecretId"] = args.secret
        try:
            body = fetch(endpoint, params)
        except Exception as exc:  # network / HTTP level
            print(f"{endpoint:<16} FAILED  {exc}", file=sys.stderr)
            failures += 1
            continue

        if isinstance(body, dict) and "error" in body:
            print(f"{endpoint:<16} ERROR   {body['error'].get('message')}", file=sys.stderr)
            failures += 1
            continue

        path = outdir / f"{endpoint}.json"
        path.write_text(json.dumps(body, indent=2))
        size = len(json.dumps(body))
        print(f"{endpoint:<16} ok      {size:>9,} bytes -> {path}")

    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
