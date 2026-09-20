"""Offseason research notes for the mandatory-drop candidates and the prospects.

Gathered from news searches on 2026-09-12, ahead of the Sep 20 draft.
Keyed by Fantrax player name. Keep short - these render in one cell.
"""
NOTES = {
    "Karel Vejmelka":
        "Healthy, no designation. Led NHL with 64 GP in 25/26; workload may ease slightly. "
        "Signed 4 more years. Sergachev/Marino added to UTA blue line.",
    "Zach Werenski":
        "Healthy. Missed time in 25/26 (lower body Dec-Jan, illness in March) but elite when "
        "playing - 40 pts in 35 games at one stretch.",
    "Tom Wilson":
        "Injury-interrupted 25/26: missed 8+ straight (lower body), foot injury blocking a shot "
        "in April. Still led WSH with 49 pts in 50 games. Olympic silver with Canada.",
    "Sebastian Aho":
        "No injury flags. Steady top-line usage in CAR.",
    "Nick Schmaltz":
        "Signed 8yr/$64M extension. First line with Keller and Peterka. RISK: has not played "
        "more than 70 games in a season since 2017-18.",
    "Jackson LaCombe":
        "New 8yr/$72M deal starts 26/27 - biggest contract in ANA history. Alternate captain, "
        "Olympic call-up. No injury flags.",
    "Seth Jarvis":
        "SHOULDER SURGERY (torn labrum + rotator cuff). Out 4-6 months; misses the start of "
        "26/27, earliest return late October, worst case midseason. CAR expects an open spot.",
    "Josh Morrissey":
        "Healthy ahead of camp. Playoff knee injury needed no surgery; normal offseason. "
        "Top pair with DeMelo. Team Canada orientation camp.",
    "Morgan Geekie":
        "Career year: team-high 39G, 68 pts, second straight 30+ goal season. Team-friendly "
        "6yr/$5.5M. Top-line or own line with PP usage. No injury flags.",
    "Mackenzie Blackwood":
        "TIMESHARE: COL enters 26/27 splitting starts with Scott Wedgewood. Lower-body surgery "
        "history (May 2025), IR stint in 25/26. Caps his win/save volume.",
    "Matvei Michkov":
        "Down year - 20G/51 pts after 26G/63 as a rookie. Reported leaner, trained with Panarin; "
        "PHI plans to integrate him more. Make-or-break framing. No current injury.",
    "Will Cuylle":
        "Contract year, RFA after 26/27. RISK: NYR winger crunch (Dorofeyev, Bjorkstrand, "
        "Perreault) could push him to the third line and cut usage.",
    "Louis Crevier":
        "Traded CHI -> BUF for Byram. Expected top-four shutdown role. Banger profile: 124 hits "
        "and 95 blocks in 25/26, both 2nd on CHI. $900K.",
    "Joseph Woll":
        "Traded TOR -> PHI for Ersson. Fresh start after a personal leave of absence in 25/26; "
        "faces a starter competition.",
    "Patrick Kane":
        "Age 37. Signed 2yr/$8M to return to CHI after three years in DET (57 pts in 67 games "
        "last season). Full no-move clause.",
    "Jamie Drysdale":
        "Offseason sports hernia surgery, expected fully recovered for camp. Was on IR in Jan "
        "(upper body). RISK: PP time not guaranteed with York and Sanheim ahead.",
    "Matthew Tkachuk":
        "Explains the low 25/26 total - surgery Aug 2025 (torn adductor + sports hernia), did "
        "not debut until January, then played the Olympics. Healthy going into 26/27.",
    "Mason Lohrei":
        "Drawing trade interest ahead of BOS camp; crowded left side, make-or-break year. "
        "Improved from -43 to +17. RFA with arbitration after 26/27.",
    "Bobby Brink":
        "Traded PHI -> MIN (Mar 2026). Upper-body injuries after the trade, then a regular "
        "healthy scratch in the playoffs. Competing for a middle-six role.",
    "Arseny Gritsyuk":
        "Solid rookie year at 25 - 13G/31 pts in 66 games, strong 5v5. Signed 3yr/$9.75M. "
        "Projected third line; NJ added Mantha and Evangelista. Late-season upper-body injury, "
        "expected ready.",
    "Victor Olofsson":
        "Signed 1yr with VGK as a UFA. Competing for a bottom-six role and 2nd PP unit - not "
        "guaranteed a lineup spot.",
    "Dalibor Dvorsky":
        "Competing for 2C in STL against Suter, McTavish and McMichael. Realistic outcome is a "
        "scoring third-line role with PP time. Breakout candidate.",
    "Cole Hutson":
        "Rookie. 10 pts in 14 games after signing out of Boston University. Top CAP defence "
        "prospect; rookie camp Sep 13-15 and likely to force a lineup spot.",

    # ---------------------------------------------------------------- prospects
    # Researched 2026-09-13 from camp coverage. These are the 23 players in the
    # MINORS block: exempt from the four drops, but they cost cap and a roster
    # spot the moment they are dressed, and they can no longer be sent back down
    # once they pass 41 GP (F/D) or 26 GP (G).
    "Sergei Murashov":
        "Likeliest scenario is he makes PIT out of camp as the 1B behind Silovs - 30-35+ "
        "starts. AHL 25/26 as the 1A: 24 W, .919, 2.20. Waiver exempt. Passes 26 GP and "
        "loses prospect status around midseason.",
    "Roman Kantserov":
        "UNDERPRICED at 50.9. Blashill has him in the top six with PP1 bumper minutes "
        "between Bedard and Frondell. Led the KHL with 36G, 64 pts in 63 GP - the most ever "
        "by a U22. ELC starts 26/27. Bedard's shoulder surgery may delay the pairing.",
    "Ilya Protas":
        "AHL Rookie of the Year with Hershey, recalled in April. 6'6/225 at 20. Rookie camp "
        "Sep 13-15, main camp Sep 17. Hershey is the likely opening assignment.",
    "Carter Yakemchuk":
        "Coin flip. Straight to main camp for a head-to-head with Nikolas Matinpalo for the "
        "bottom-pair RD spot; THN gives Matinpalo the edge. NHL.com's projected lineup has "
        "him on the third pair. Loses the duel and he is back in Belleville.",
    "Harrison Brunicke":
        "Best pedigree, worst timing. PIT's No. 1 prospect (THN), No. 2 in PensBurgh's "
        "25-under-25, Canada at the 2026 WJC. Blocked on the right side by Karlsson, Letang, "
        "Korczak and van Riemsdyk - AHL expected.",
    "Samuel Honzek":
        "UNDERPRICED at 16.6. Projected on Calgary's shutdown line with Backlund and Farabee; "
        "was on an NHL roster 50+ games in 25/26 and would have to play his way OFF the team. "
        "RISK: both of his last two seasons ended early on injury (18 GP a year ago). Low "
        "fantasy ceiling - no PP time.",
    "Maveric Lamoureux":
        "Real runway. UTA lost Durzi and Cole; Lamoureux and Simashev are named as the "
        "replacements. 2022 1st-rounder (29th). 14 pts in 42 AHL games.",
    "Rutger McGroarty":
        "Bubble. In the mix out of camp but competing with Avery Hayes and Filip Hallander. "
        "34 pts in 30 AHL games, but only 6 pts in 24 NHL games at 11:58 TOI.",
    "William Stromgren":
        "NHL debut Jan 2026, 3 games. Projected for 10-20 games this season. Has fallen "
        "outside CGY's top 20 prospects as the pipeline deepened.",
    "Hunter Brzustewicz":
        "Starts in the AHL - CGY's blue line is crowded with Hanley and Pachal ahead. 26/27 "
        "is the final year of his ELC and his final year of waiver exemption.",
    "Kasper Halttunen":
        "Acquired from SJS in the June William Eklund trade. 6'4/205 power forward with an "
        "excellent shot; 35 pts in 69 AHL games. Could push for a look but Belleville for "
        "most of the year.",
    "Quentin Musty":
        "2023 26th overall. Strong first AHL year at 20 - 35 pts in 53 games with the "
        "Barracuda. Finisher who projects middle-six. Another AHL season expected.",
    "Terik Parascak":
        "2024 17th overall. 79 pts in 61 WHL games, re-assigned to Hershey in April 2026. "
        "RISK: skating and pace are the knocks. A middle-six contributor eventually, not now.",
    "Henry Brzustewicz":
        "KEEPER ASSET. NHL.com ranks him the Kings' No. 1 prospect. 2025 1st-rounder (31st), "
        "54 pts in 59 OHL games with London. ELC starts 26/27. Could land in the NHL, AHL or "
        "OHL this season.",
    "Alfons Freij":
        "2024 37th overall. Coming to North America a year early after development camp talks. "
        "Competes at camp but likely opens with the Moose. Gold with Sweden at the 2026 WJC.",
    "Joel Blomqvist":
        "No. 3 on PIT's depth chart behind Silovs and Murashov - redundant with Murashov on "
        "this roster. 16-5-4, .913, 2.40 in the AHL; no NHL games in 25/26. Signed 2yr/$875K.",
    "Dylan Garand":
        "Sleeper. Hartford starter knocking on the door - 2-0-1 in three NHL appearances in "
        "25/26. Jonathan Quick turns 40 in January and Garand is the in-house replacement.",
    "James Malatesta":
        "One-year two-way deal ($850K NHL / $95K AHL). 13 career NHL games over two seasons, "
        "none in 25/26. 18 pts and a team-high 87 PIM in 57 AHL games. Critical year.",
    "Max Plante":
        "2026 HOBEY BAKER WINNER - 52 pts in 40 games at UMD. Back on UMD's roster for 26/27, "
        "so he is a two-year hold, not a this-year asset. Drafted 47th in 2024.",
    "Yegor Sidorov":
        "FADING. Clear step back in his second AHL year - .43 pts/game in 68 games after 164 "
        "pts in 119 WHL games. Skilled but one-dimensional. RFA after 26/27.",
    "Julian Lutz":
        "FADING. Two goals in 24 AHL games in 25/26; leaned into a physical role instead. "
        "RFA after 26/27.",
    "Brett Harrison":
        "DEAD SLOT. Contract expired after 25/26 - RFA with PHI, currently unsigned. Under the "
        "holdout rule he takes a roster spot and his cap hit applies the moment he signs.",
    "Tristan Lennox":
        "DEAD SLOT. Contract expired after 25/26 and he is an unrestricted free agent with no "
        "NHL team. Dropping him clears a prospect slot for the 5-round prospect draft.",
}
