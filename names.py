"""Matching one player's name across sources that spell it differently.

Fantrax, the NHL and DobberProspects each have their own idea of a player's
name, and the disagreements are systematic rather than random: a short form
(Zach/Zachary, Gabe/Gabriel), a nickname that is not a prefix (Mikey/Michael,
Joe/Joseph, Jack/John), or a transliteration (Andrei/Andrey, Alexei/Alexey,
Scherbakov/Shcherbakov).

The rule everything here serves: compare the first and last names separately.
Scoring a whole name in one go is what pairs Dylan Roobroeck with his brother
Ryan and Michael Hagens with Michael Hage - the surnames carry most of the
characters, so a wrong first name barely moves the score.
"""
import difflib
import re
import unicodedata

# Every spelling of a given name collapsed onto one canonical key. This has
# to be a table rather than a prefix test: "Zach" really is "Zachary", but
# "Jack" is NOT "Jackson" and "Jason" is NOT "Jackson", and a prefix or
# similarity test cannot tell those apart - it paired Jack Smith AND Jason
# Smith with Jackson Smith's scouting report.
_GROUPS = [
    ("alex", "alexander alexandre aleksander aleksandr sasha"),
    ("michael", "mike mikey mick"),
    ("john", "jack johnny jonathan johnathan jon"),
    ("robert", "bob bobby rob robbie"),
    ("william", "will billy bill liam"),
    ("anthony", "tony"),
    ("nicholas", "nick nicky nikolas"),
    ("daniel", "dan danny"),
    ("thomas", "tom tommy"),
    ("jacob", "jake"),
    ("james", "jim jimmy jamie"),
    ("joseph", "joe joey"),
    ("patrick", "pat"),
    ("richard", "rick ricky dick"),
    ("steven", "steve stephen"),
    ("theodore", "ted teddy"),
    ("timothy", "tim"),
    ("vincent", "vinny"),
    ("zachary", "zach zack"),
    ("gabriel", "gabe"),
    ("christopher", "chris"),
    ("matthew", "matt"),
    ("mitchell", "mitch"),
    ("andrew", "andy drew"),
    ("charles", "charlie chuck"),
    ("benjamin", "ben benji"),
    ("samuel", "sam sammy"),
    ("gregory", "greg"),
    ("jeffrey", "jeff"),
    ("joshua", "josh"),
    ("edward", "ed eddie"),
    ("peter", "pete"),
]
CANON = {}
for _c, _alts in _GROUPS:
    CANON[_c] = _c
    for _a in _alts.split():
        CANON[_a] = _c

SUFFIX = {"jr", "sr", "ii", "iii", "iv"}


def norm(s):
    """A name down to bare letters: no accents, no punctuation, no case."""
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^a-z]", "", s.lower())


def parts(name):
    """A name or a slug -> (first, last), split the same way for both so the
    two can be compared token by token. Hyphens separate, so "Jean-Cristoph
    Lemieux" and "jean-christoph-lemieux" both give ("jean", "lemieux"); a
    Jr/Sr suffix is dropped before the last token is taken."""
    toks = [norm(t) for t in re.split(r"[\s-]+", str(name)) if norm(t)]
    while len(toks) > 1 and toks[-1] in SUFFIX:
        toks.pop()
    return (toks[0], toks[-1]) if toks else ("", "")


def ratio(a, b):
    return difflib.SequenceMatcher(None, a, b).ratio()


def same_last(a, b, cutoff=0.88):
    """Surnames of one man. Strict: it is the half of the name that
    distinguishes people, so only a spelling variant or a transposition may
    differ (Talliefer/Taillefer, 0.89). Hagens/Hage scores 0.80 and is
    correctly a different player, which is what sets the floor."""
    return a == b or ratio(a, b) >= cutoff


def same_first(a, b, cutoff=0.80):
    """First names that plausibly belong to one man: the same name, a known
    short form, or a transliteration.

    Similarity alone is not enough - Bryan/Ryan score 0.89 and Jason/Jackson
    0.83, and neither is the same man. So a fuzzy pass also has to be nearly
    the same length, which is true of transliterations (Egor/Yegor,
    Andrei/Andrey, Artyom/Artyon) and false of a different name that happens
    to share letters."""
    if a == b or CANON.get(a, a) == CANON.get(b, b):
        return True
    return ratio(a, b) >= cutoff and abs(len(a) - len(b)) <= 1


def same_person(a, b):
    """Two full names that refer to the same player."""
    af, al = parts(a)
    bf, bl = parts(b)
    return same_last(al, bl) and same_first(af, bf)
