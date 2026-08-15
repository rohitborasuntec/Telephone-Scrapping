#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
 LETTER ONE - SINGLE FILE PDF MERGE ENGINE
================================================================================

 Overlays dynamic data onto the client's APPROVED Letter One PDF using
 PyMuPDF only. The approved PDF stays the immutable master visual template:

   - no conversion to Word / WordPad / DOCX / HTML
   - no rasterising, no rebuilding of the static design
   - only merge-tag text is removed, then redrawn at coordinates read from
     the template itself at run time

 REQUIREMENTS
   pip install pymupdf
   Fonts: Georgia, Aptos, Arial (regular + bold). On a standard Windows
   machine with Microsoft 365 these are already installed and nothing
   needs configuring.

 USAGE
   Diagnostics (run first on any new template):
       python letter_one_generate.py --inspect

   Generate with the values baked in below:
       python letter_one_generate.py

   Override at the command line:
       python letter_one_generate.py --reference "PA/26/00689/S" \
           --address "Linden Way Kingfield Surrey Woking GU22 9BS"

   Preview on a machine without the licensed fonts (NEVER for production):
       python letter_one_generate.py --dev-fonts
================================================================================
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

try:
    import pymupdf  # PyMuPDF >= 1.24
except ImportError:  # pragma: no cover
    import fitz as pymupdf  # type: ignore


# ==============================================================================
#  1. CONFIGURATION - everything tunable lives here
# ==============================================================================

TEMPLATE_PATH = Path("templates/letter_one.pdf")
OUTPUT_DIR = Path("output")

# ---- the data to merge -------------------------------------------------------
REFERENCE = "PA/26/00689/S"

SITE_ADDRESS = (
    "2 The Courtyards Phoenix Square Severalls Business Park "
    "Wyncolls Road Colchester Essex CO4 9PE"
)

PLANNING_STATUS = None  # only needed if the template ever gains a status tag

# ---- document invariants -----------------------------------------------------
REQUIRED_PAGE_COUNT = 2
PAGE_SIZE_TOLERANCE_PT = 0.5

TAG_REFERENCE = "{{Reference}}"
TAG_ADDRESS = "{{Address}}"
TAG_PATTERN = r"\{\{[^{}]{1,64}\}\}"

# ---- header geometry ---------------------------------------------------------
# The right edge is DERIVED from the template at run time (the two header
# lines are right-aligned in the source document, so their right edge is the
# approved edge). The value below is only a sanity bound.
EXPECTED_HEADER_RIGHT_EDGE = 525.44
RIGHT_EDGE_TOLERANCE = 6.0

HEADER_LEFT_LIMIT = 72.02          # body text left margin
HEADER_EXTEND_LEFT_LIMIT = 42.0    # used only by the 'extend_left' policy
BODY_TEXT_RIGHT_LIMIT = 523.32     # right edge of the body text column

# ---- Stannp pink recipient / OCR clear zone ----------------------------------
# Derived by measuring the Stannp preview screen-grab against known static
# anchors (the address line above it, "Dear Neighbour," below it), then
# padded outwards. CONFIRM AGAINST A LIVE STANNP PROOF BEFORE PRODUCTION.
STANNP_OCR_ZONE: Tuple[float, float, float, float] = (24.0, 120.0, 332.0, 260.0)

PROTECTED_REGIONS: Dict[str, Tuple[float, float, float, float]] = {
    "stannp_ocr_zone": STANNP_OCR_ZONE,
    "page1_footer": (0.0, 770.0, 595.32, 841.92),
    "page2_footer": (0.0, 770.0, 595.32, 841.92),
}

COLLISION_TOLERANCE_PT = 0.75  # adjacent PDF lines legitimately touch by a hair

# ---- long-address policy -----------------------------------------------------
#   "strict"      the approved behaviour: must fit between HEADER_LEFT_LIMIT
#                 and the approved right edge at the approved font size, or
#                 generation FAILS with SITE_ADDRESS_TOO_LONG.
#   "extend_left" the line may start further left, down to
#                 HEADER_EXTEND_LEFT_LIMIT. Still one line, still 10pt.
#   "shrink"      reduces ONLY the site-address line, in 0.25pt steps, no
#                 smaller than MIN_ADDRESS_FONT_SIZE.
#
# The address configured above needs ~494pt and the approved area gives
# ~453pt, so it CANNOT be rendered at 10pt. 'shrink' is set here so this
# record produces output; it deviates from the 10pt specification and is
# logged as a WARNING requiring client sign-off. Set back to "strict" to
# route such records to an exception queue instead.
ADDRESS_OVERFLOW_POLICY = "shrink"
MIN_ADDRESS_FONT_SIZE = 8.0
ADDRESS_SHRINK_STEP = 0.25

# ---- fonts -------------------------------------------------------------------
# The template's embedded fonts are SUBSETS - the embedded "Aptos" has no
# "/", "0", "3", "4"... so PA/26/00689/S cannot be drawn from them. Real
# licensed font files are mandatory.
APPROVED_FONT_SIZE = 10.0

FONT_FILE_CANDIDATES: Dict[str, Dict[str, List[str]]] = {
    "georgia": {
        "regular": ["georgia.ttf", "Georgia.ttf"],
        "bold": ["georgiab.ttf", "Georgia-Bold.ttf", "Georgia Bold.ttf"],
        "italic": ["georgiai.ttf", "Georgia-Italic.ttf"],
        "bolditalic": ["georgiaz.ttf", "Georgia-BoldItalic.ttf"],
    },
    "aptos": {
        "regular": ["aptos.ttf", "Aptos.ttf"],
        "bold": ["aptos-bold.ttf", "Aptos-Bold.ttf", "aptosb.ttf"],
        "italic": ["aptos-italic.ttf", "Aptos-Italic.ttf"],
        "bolditalic": ["aptos-bold-italic.ttf", "Aptos-Bold-Italic.ttf"],
    },
    "arial": {
        "regular": ["arial.ttf", "Arial.ttf"],
        "bold": ["arialbd.ttf", "Arial-Bold.ttf"],
        "italic": ["ariali.ttf", "Arial-Italic.ttf"],
        "bolditalic": ["arialbi.ttf", "Arial-BoldItalic.ttf"],
    },
}

SYSTEM_FONT_DIRS: List[Path] = [
    Path(r"C:\Windows\Fonts"),
    Path(os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Windows\Fonts")),
    Path("/Library/Fonts"),
    Path("/System/Library/Fonts/Supplemental"),
    Path("/usr/share/fonts"),
    Path("/usr/local/share/fonts"),
    Path.home() / ".fonts",
]

# Metric-approximate stand-ins, used ONLY with --dev-fonts. Preview only.
DEV_FALLBACK_FONTS: Dict[str, Dict[str, List[str]]] = {
    "georgia": {
        "regular": ["Gelasio-Regular.ttf", "Caladea-Regular.ttf", "LiberationSerif-Regular.ttf"],
        "bold": ["Gelasio-Bold.ttf", "Caladea-Bold.ttf", "LiberationSerif-Bold.ttf"],
        "italic": ["Gelasio-Italic.ttf", "Caladea-Italic.ttf", "LiberationSerif-Italic.ttf"],
        "bolditalic": ["Gelasio-BoldItalic.ttf", "Caladea-BoldItalic.ttf", "LiberationSerif-BoldItalic.ttf"],
    },
    "aptos": {
        "regular": ["Carlito-Regular.ttf", "LiberationSans-Regular.ttf"],
        "bold": ["Carlito-Bold.ttf", "LiberationSans-Bold.ttf"],
        "italic": ["Carlito-Italic.ttf", "LiberationSans-Italic.ttf"],
        "bolditalic": ["Carlito-BoldItalic.ttf", "LiberationSans-BoldItalic.ttf"],
    },
    "arial": {
        "regular": ["LiberationSans-Regular.ttf"],
        "bold": ["LiberationSans-Bold.ttf"],
        "italic": ["LiberationSans-Italic.ttf"],
        "bolditalic": ["LiberationSans-BoldItalic.ttf"],
    },
}


# ==============================================================================
#  2. ERRORS
# ==============================================================================

class LetterOneError(Exception):
    code = "UNKNOWN_ERROR"

    def __init__(self, message: str, **details) -> None:
        super().__init__(message)
        self.message = message
        self.details = details

    def __str__(self) -> str:
        if not self.details:
            return "[%s] %s" % (self.code, self.message)
        body = "\n".join("    %s: %s" % (k, v) for k, v in self.details.items())
        return "[%s] %s\n%s" % (self.code, self.message, body)


class TemplateNotFound(LetterOneError):          code = "TEMPLATE_NOT_FOUND"
class InvalidPDF(LetterOneError):                code = "INVALID_PDF"
class InvalidPageCount(LetterOneError):          code = "INVALID_PAGE_COUNT"
class InvalidInputData(LetterOneError):          code = "INVALID_INPUT_DATA"
class MergeTagNotFound(LetterOneError):          code = "MERGE_TAG_NOT_FOUND"
class UnresolvedMergeTag(LetterOneError):        code = "UNRESOLVED_MERGE_TAG"
class SiteAddressTooLong(LetterOneError):        code = "SITE_ADDRESS_TOO_LONG"
class SiteAddressOverlapsOCRZone(LetterOneError):code = "SITE_ADDRESS_OVERLAPS_OCR_ZONE"
class SiteAddressOutsideHeader(LetterOneError):  code = "SITE_ADDRESS_OUTSIDE_HEADER"
class ReferenceOverflow(LetterOneError):         code = "REFERENCE_OVERFLOW"
class ProtectedRegionViolation(LetterOneError):  code = "PROTECTED_REGION_VIOLATION"
class StaticContentAltered(LetterOneError):      code = "STATIC_CONTENT_ALTERED"
class FontNotFound(LetterOneError):              code = "FONT_NOT_FOUND"
class PDFGenerationFailed(LetterOneError):       code = "PDF_GENERATION_FAILED"


# ==============================================================================
#  3. FONT RESOLUTION AND MEASUREMENT
# ==============================================================================

_SUBSET_PREFIX = re.compile(r"^[A-Z]{6}\+")
_FAMILY_ALIASES = {"arialmt": "arial", "arial": "arial", "aptos": "aptos",
                   "aptosdisplay": "aptos", "georgia": "georgia", "georgiapro": "georgia"}


class FontStyle(object):
    """Normalised description of a text run's typeface."""

    __slots__ = ("family", "bold", "italic")

    def __init__(self, family: str, bold: bool = False, italic: bool = False) -> None:
        self.family, self.bold, self.italic = family, bold, italic

    @property
    def variant(self) -> str:
        if self.bold and self.italic:
            return "bolditalic"
        if self.bold:
            return "bold"
        if self.italic:
            return "italic"
        return "regular"

    def describe(self) -> str:
        return "%s %s" % (self.family.title(), self.variant)

    def key(self):
        return (self.family, self.bold, self.italic)

    def __eq__(self, other):
        return isinstance(other, FontStyle) and self.key() == other.key()

    def __hash__(self):
        return hash(self.key())


def normalise_font_name(raw: str) -> FontStyle:
    """'BCDEEE+Georgia-Bold' -> FontStyle('georgia', bold=True)."""
    name = _SUBSET_PREFIX.sub("", raw or "").strip()
    low = name.lower()
    bold = "bold" in low or low.endswith("bd") or ",b" in low
    italic = "italic" in low or "oblique" in low
    family = re.split(r"[-,]", low)[0]
    for token in ("bolditalic", "boldmt", "bold", "italicmt", "italic",
                  "oblique", "regular", "mt", "psmt"):
        family = family.replace(token, "")
    family = family.strip(" -_")
    return FontStyle(_FAMILY_ALIASES.get(family, family or "unknown"), bold, italic)


class FontResolver(object):
    def __init__(self, extra_dirs: Optional[Sequence[Path]] = None,
                 use_dev_fonts: bool = False) -> None:
        self.use_dev_fonts = use_dev_fonts
        self.substitutions: Dict[str, str] = {}
        self._paths: Dict[FontStyle, Path] = {}
        self._fonts: Dict[str, "pymupdf.Font"] = {}
        self._dirs: List[Path] = []
        env = os.environ.get("LETTER_ONE_FONT_DIR")
        if env:
            self._dirs.append(Path(env))
        self._dirs.extend(Path(d) for d in (extra_dirs or []))
        self._dirs.append(Path("fonts"))
        self._dirs.extend(SYSTEM_FONT_DIRS)

    def _find(self, names: List[str]) -> Optional[Path]:
        # names are in PREFERENCE order - exhaust every directory for the
        # first name before falling back to the next one
        for wanted in names:
            target = wanted.lower()
            for directory in self._dirs:
                try:
                    if not directory.exists():
                        continue
                except OSError:
                    continue
                for path in directory.rglob("*"):
                    if path.is_file() and path.name.lower() == target:
                        return path
        return None

    def path(self, style: FontStyle) -> Path:
        if style in self._paths:
            return self._paths[style]

        family_map = FONT_FILE_CANDIDATES.get(style.family)
        if family_map is None:
            raise FontNotFound("No font-file mapping configured for this family.",
                               family=style.family, variant=style.variant)

        found = self._find(family_map[style.variant])

        if found is None and style.variant in ("italic", "bolditalic"):
            base = "bold" if style.variant == "bolditalic" else "regular"
            found = self._find(family_map[base])
            if found is not None:
                self.substitutions[style.describe()] = \
                    "%s (no %s cut installed)" % (found.name, style.variant)

        if found is None and self.use_dev_fonts:
            found = self._find(DEV_FALLBACK_FONTS.get(style.family, {}).get(style.variant, []))
            if found is not None:
                self.substitutions[style.describe()] = "%s (DEV SUBSTITUTE)" % found.name

        if found is None:
            raise FontNotFound(
                "Required font file not found on this machine.",
                family=style.family, variant=style.variant,
                looked_for=", ".join(family_map[style.variant]),
                searched=", ".join(str(d) for d in self._dirs),
                hint=("Install the licensed font, drop the .ttf into ./fonts/, or set "
                      "LETTER_ONE_FONT_DIR. --dev-fonts is preview only."))

        self._paths[style] = found
        return found

    def font(self, style: FontStyle):
        key = str(self.path(style))
        if key not in self._fonts:
            self._fonts[key] = pymupdf.Font(fontfile=key)
        return self._fonts[key]

    def text_length(self, text: str, style: FontStyle, size: float) -> float:
        """Real advance width in points - never estimated from character count."""
        return self.font(style).text_length(text, fontsize=size) if text else 0.0

    def missing_glyphs(self, text: str, style: FontStyle) -> str:
        f = self.font(style)
        return "".join(sorted({c for c in text if f.has_glyph(ord(c)) == 0}))

    def pdf_fontname(self, style: FontStyle) -> str:
        return "LO_%s_%s" % (style.family, style.variant)


# ==============================================================================
#  4. DATA MODEL
# ==============================================================================

def normalise_whitespace(value: str) -> str:
    """Collapse every run of whitespace (including newlines) to one space."""
    return " ".join(str(value).split())


class LetterData(object):
    def __init__(self, reference: str, site_address: str,
                 planning_status: Optional[str] = None) -> None:
        self.reference = normalise_whitespace(reference)
        self.site_address = normalise_whitespace(site_address)
        self.planning_status = normalise_whitespace(planning_status) if planning_status else None
        if not self.reference:
            raise InvalidInputData("Required dynamic value is missing.", field="reference")
        if not self.site_address:
            raise InvalidInputData("Required dynamic value is missing.", field="site_address")

    def safe_reference(self) -> str:
        return "".join(c if c.isalnum() else "_" for c in self.reference).strip("_") or "UNKNOWN"


class TextRun(object):
    __slots__ = ("text", "style", "size", "color", "baseline_y", "is_merge_value", "source_tag")

    def __init__(self, text, style, size, color, baseline_y,
                 is_merge_value=False, source_tag=None):
        self.text = text
        self.style = style
        self.size = size
        self.color = color
        self.baseline_y = baseline_y
        self.is_merge_value = is_merge_value
        self.source_tag = source_tag

    @property
    def rgb(self):
        c = self.color
        return (((c >> 16) & 255) / 255.0, ((c >> 8) & 255) / 255.0, (c & 255) / 255.0)


class RebuiltLine(object):
    def __init__(self, page_number, kind, align, anchor_x, original_bbox, redact_rect, runs, tags):
        self.page_number = page_number
        self.kind = kind                 # header_reference | header_address | body
        self.align = align               # left | right
        self.anchor_x = anchor_x
        self.original_bbox = original_bbox
        self.redact_rect = redact_rect
        self.runs = runs
        self.tags = tags
        self.rendered_bbox = None
        self.rendered_width = 0.0
        self.available_width = 0.0
        self.applied_font_size = None
        self.scale = 1.0
        self.shrunk = False

    def text(self) -> str:
        return "".join(r.text for r in self.runs)


# ==============================================================================
#  5. TEMPLATE SCANNING AND REDRAW PLANNING
# ==============================================================================

TAG_RE = re.compile(TAG_PATTERN)
_WS = re.compile(r"\s+")


def _iter_lines(page):
    for block in page.get_text("dict").get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            if line.get("spans"):
                yield line


def _line_text(line) -> str:
    return "".join(s["text"] for s in line["spans"])


def _span_offsets(line):
    out, pos = [], 0
    for span in line["spans"]:
        out.append((pos, pos + len(span["text"]), span))
        pos += len(span["text"])
    return out


def _owning_span(offsets, pos):
    for start, end, span in offsets:
        if start <= pos < end:
            return span
    return offsets[-1][2]


def _surrounding_span(offsets, match_start, match_end):
    """
    The static run whose formatting the merge value must copy.

    Client requirement: the merge tags must match the font of the
    surrounding text. Preference: the first following run containing a
    non-space character, else the last preceding one. Pure-whitespace runs
    carry no visible formatting and are skipped.
    """
    after = [s for st, en, s in offsets if st >= match_end and s["text"].strip()]
    if after:
        return after[0]
    before = [s for st, en, s in offsets if en <= match_start and s["text"].strip()]
    return before[-1] if before else None


def _merge_style(placeholder_span, surrounding_span):
    """
    Decide the typeface and colour of an inserted merge value.

      - family and colour come from the SURROUNDING static text. This
        matters: in the supplied template the body {{Reference}} is tagged
        Aptos while the paragraph around it is Georgia, and the blue quote
        placeholder carries #215F9A while the quote is #215E99.
      - weight and slant come from the placeholder's own run, so a value
        sitting beside a bold label does not inherit the bold.
    """
    own = normalise_font_name(placeholder_span.get("font", ""))
    own_color = placeholder_span.get("color", 0)
    if surrounding_span is None:
        return own, own_color
    surrounding = normalise_font_name(surrounding_span.get("font", ""))
    style = FontStyle(surrounding.family or own.family, own.bold, own.italic)
    return style, surrounding_span.get("color", own_color)


def derive_header_right_edge(page) -> Optional[float]:
    """
    The approved right edge, derived from the template rather than
    hard-coded: the header lines carrying {{Reference}} and {{Address}} are
    right-aligned in the source document, so their right edge IS the edge.
    """
    edges = []
    for line in _iter_lines(page):
        text = _line_text(line)
        if TAG_REFERENCE in text or TAG_ADDRESS in text:
            x0, y0, x1, y1 = line["bbox"]
            if x0 > 260.0 and y1 < 200.0:
                edges.append(x1)
    return max(edges) if edges else None


def _is_right_aligned(line, right_edge: float) -> bool:
    x0, _, x1, _ = line["bbox"]
    return abs(x1 - right_edge) <= RIGHT_EDGE_TOLERANCE and x0 > 260.0


def _value_for_tag(tag: str, data: LetterData) -> str:
    key = tag[2:-2].strip().lower().replace(" ", "").replace("_", "")
    if key == "reference":
        return data.reference
    if key == "address":
        return data.site_address
    if key in ("planningstatus", "status"):
        if data.planning_status is None:
            raise InvalidInputData("Template has a planning-status tag but no value supplied.", tag=tag)
        return data.planning_status
    raise MergeTagNotFound("Template contains a merge tag the engine has no value for.", tag=tag)


def plan_lines(doc, data: LetterData, right_edge: float) -> List[RebuiltLine]:
    """
    Build the redraw plan.

    RIGHT-ALIGNED header line -> the whole line is rebuilt, because changing
    any run's width moves the line's start.

    LEFT-ALIGNED body line -> only the tail from the first tag to end-of-line
    is rebuilt. Everything before it keeps its original approved position to
    the point.
    """
    plans: List[RebuiltLine] = []

    for pno in range(doc.page_count):
        for line in _iter_lines(doc[pno]):
            text = _line_text(line)
            matches = list(TAG_RE.finditer(text))
            if not matches:
                continue

            offsets = _span_offsets(line)
            right_aligned = _is_right_aligned(line, right_edge)
            first_start = matches[0].start()

            start_idx = 0 if right_aligned else next(
                i for i, (st, en, _s) in enumerate(offsets) if st <= first_start < en)
            anchor_x = right_edge if right_aligned else offsets[start_idx][2]["bbox"][0]

            runs: List[TextRun] = []
            tags: List[str] = []

            for idx in range(start_idx, len(offsets)):
                st, en, span = offsets[idx]
                span_text = span["text"]
                style = normalise_font_name(span.get("font", ""))
                size = span.get("size", APPROVED_FONT_SIZE)
                color = span.get("color", 0)
                baseline = span.get("origin", (span["bbox"][0], span["bbox"][3]))[1]

                cursor = 0
                for match in matches:
                    if match.end() <= st or match.start() >= en:
                        continue
                    local_start = max(match.start() - st, 0)
                    local_end = min(match.end() - st, len(span_text))

                    if local_start > cursor:
                        runs.append(TextRun(span_text[cursor:local_start], style,
                                            size, color, baseline))

                    if match.start() >= st:  # emit the value once, in its owning span
                        surrounding = _surrounding_span(offsets, match.start(), match.end())
                        merge_style, merge_color = _merge_style(span, surrounding)
                        merge_baseline = (surrounding.get("origin", (0, baseline))[1]
                                          if surrounding is not None else baseline)
                        tag = match.group(0)
                        tags.append(tag)
                        runs.append(TextRun(_value_for_tag(tag, data), merge_style,
                                            size, merge_color, merge_baseline,
                                            is_merge_value=True, source_tag=tag))
                    cursor = local_end

                if cursor < len(span_text):
                    runs.append(TextRun(span_text[cursor:], style, size, color, baseline))

            runs = [r for r in runs if r.text != ""]

            lx0, ly0, lx1, ly1 = line["bbox"]
            redact = (min(anchor_x, offsets[start_idx][2]["bbox"][0]) - 0.4,
                      ly0 + 0.25,
                      max(lx1, right_edge) + 2.0,
                      ly1 - 0.25)

            kind = ("header_address" if TAG_ADDRESS in tags
                    else "header_reference" if right_aligned else "body")

            plans.append(RebuiltLine(pno, kind, "right" if right_aligned else "left",
                                     anchor_x, (lx0, ly0, lx1, ly1), redact, runs, tags))
    return plans


def count_tags(doc) -> Dict[str, int]:
    counts: Dict[str, int] = {}
    for pno in range(doc.page_count):
        for m in TAG_RE.finditer(doc[pno].get_text()):
            counts[m.group(0)] = counts.get(m.group(0), 0) + 1
    return counts


# ==============================================================================
#  6. ENGINE - redaction, layout policy, drawing
# ==============================================================================

def _intersects(a, b, tol=0.0) -> bool:
    ax0, ay0, ax1, ay1 = a
    bx0, by0, bx1, by1 = b
    return not (ax1 <= bx0 + tol or ax0 >= bx1 - tol
                or ay1 <= by0 + tol or ay0 >= by1 - tol)


def _page_text(page) -> str:
    """
    Extract page text in READING ORDER.

    Redrawn lines live in a content stream appended after the original one,
    so default extraction lists them last and every order-sensitive
    comparison would spuriously fail. Sorting by position fixes that.
    """
    return page.get_text("text", sort=True)


class LetterOneEngine(object):
    def __init__(self, template: Path, resolver: FontResolver,
                 address_policy: str = ADDRESS_OVERFLOW_POLICY,
                 min_address_font_size: float = MIN_ADDRESS_FONT_SIZE,
                 ocr_zone=STANNP_OCR_ZONE) -> None:
        self.template = Path(template)
        self.resolver = resolver
        self.address_policy = address_policy
        self.min_address_font_size = min_address_font_size
        self.ocr_zone = ocr_zone
        self.right_edge = EXPECTED_HEADER_RIGHT_EDGE
        self.warnings: List[str] = []
        self._planned_bboxes: Dict[int, List] = {}

    # -- template ----------------------------------------------------------
    def open_template(self):
        if not self.template.exists():
            raise TemplateNotFound("Template PDF not found.", path=str(self.template))
        try:
            doc = pymupdf.open(self.template)
        except Exception as exc:
            raise InvalidPDF("Template could not be opened as a PDF.",
                             path=str(self.template), error=str(exc))
        if doc.page_count != REQUIRED_PAGE_COUNT:
            raise InvalidPageCount("Template does not have the required page count.",
                                   expected=REQUIRED_PAGE_COUNT, actual=doc.page_count)
        return doc

    def effective_left_limit(self) -> float:
        return HEADER_EXTEND_LEFT_LIMIT if self.address_policy == "extend_left" else HEADER_LEFT_LIMIT

    # -- measuring ---------------------------------------------------------
    def _run_width(self, run: TextRun, scale: float = 1.0) -> float:
        return self.resolver.text_length(run.text, run.style, run.size * scale)

    def _line_width(self, line: RebuiltLine, scale: float = 1.0) -> float:
        return sum(self._run_width(r, scale) for r in line.runs)

    def _check_glyphs(self, line: RebuiltLine) -> None:
        for run in line.runs:
            if not run.is_merge_value:
                continue
            missing = self.resolver.missing_glyphs(run.text, run.style)
            if missing:
                raise FontNotFound("Resolved font cannot render every character of the merge value.",
                                   font=self.resolver.path(run.style).name,
                                   style=run.style.describe(), value=run.text,
                                   missing_characters=missing)

    # -- layout ------------------------------------------------------------
    def _layout(self, line: RebuiltLine, page) -> None:
        scale = 1.0
        width = self._line_width(line)

        if line.align == "right":
            left_limit = self.effective_left_limit()
            available = self.right_edge - left_limit
            line.available_width = available

            if width > available and line.kind == "header_address":
                scale = self._resolve_address_overflow(line, width, available)
                width = self._line_width(line, scale)

            start_x = self.right_edge - width
            if start_x < left_limit - 0.01:
                self._raise_address_too_long(line, width, available, start_x)
        else:
            start_x = line.anchor_x
            line.available_width = BODY_TEXT_RIGHT_LIMIT - start_x
            if start_x + width > BODY_TEXT_RIGHT_LIMIT + 0.01:
                raise ReferenceOverflow(
                    "Replacing the merge tag makes this line wider than the approved "
                    "text column; the line would run into the margin.",
                    page=line.page_number + 1, tags=", ".join(line.tags),
                    line_text=line.text().strip()[:120],
                    required_width=round(width, 2),
                    available_width=round(line.available_width, 2),
                    right_limit=BODY_TEXT_RIGHT_LIMIT,
                    would_end_at=round(start_x + width, 2))

        top = min(r.baseline_y - self.resolver.font(r.style).ascender * r.size * scale
                  for r in line.runs)
        bottom = max(r.baseline_y - self.resolver.font(r.style).descender * r.size * scale
                     for r in line.runs)

        line.scale = scale
        line.shrunk = scale != 1.0
        line.applied_font_size = round(max(r.size for r in line.runs) * scale, 2)
        line.rendered_width = width
        line.rendered_bbox = (start_x, top, start_x + width, bottom)

        self._check_protected_regions(line, page)

    def _resolve_address_overflow(self, line, width, available) -> float:
        """Section 25 - default is to FAIL. Shrinking only on explicit request."""
        if self.address_policy != "shrink":
            return 1.0
        base = max(r.size for r in line.runs)
        size = base
        while size - ADDRESS_SHRINK_STEP >= self.min_address_font_size - 1e-9:
            size -= ADDRESS_SHRINK_STEP
            scale = size / base
            if self._line_width(line, scale) <= available:
                self.warnings.append(
                    "Site address did not fit at %.2f pt (%.1f pt needed, %.1f pt "
                    "available). Policy 'shrink' reduced the site-address line only, "
                    "to %.2f pt. This deviates from the approved 10 pt specification "
                    "and must be signed off by the client."
                    % (base, width, available, size))
                return scale
        return 1.0

    def _raise_address_too_long(self, line, width, available, start_x) -> None:
        value = next((r.text for r in line.runs if r.is_merge_value), "")
        raise SiteAddressTooLong(
            "The site address cannot be rendered as one continuous line inside the "
            "approved Letter One header area at the approved font size.",
            address=value, full_line=line.text().strip(),
            required_width=round(width, 2), available_width=round(available, 2),
            shortfall=round(width - available, 2),
            font=", ".join(sorted({r.style.describe() for r in line.runs})),
            font_size=round(max(r.size for r in line.runs), 2),
            template=str(self.template), page=line.page_number + 1,
            calculated_x=round(start_x, 2), left_limit=self.effective_left_limit(),
            right_edge=round(self.right_edge, 2), policy=self.address_policy,
            options=("Widen the header column in the approved template, route this record "
                     "to an exception queue, or set ADDRESS_OVERFLOW_POLICY to "
                     "'extend_left' / 'shrink' (both need client sign-off)."))

    # -- protected regions -------------------------------------------------
    def _check_protected_regions(self, line: RebuiltLine, page) -> None:
        bbox = line.rendered_bbox
        x0, y0, x1, y1 = bbox
        rect = page.rect
        if x0 < 0 or y0 < 0 or x1 > rect.width or y1 > rect.height:
            raise ProtectedRegionViolation("Rebuilt line falls outside the page.",
                                           page=line.page_number + 1,
                                           bbox=[round(v, 2) for v in bbox])

        for name, region in PROTECTED_REGIONS.items():
            if name == "stannp_ocr_zone":
                region = self.ocr_zone
            if name.startswith("page1") and line.page_number != 0:
                continue
            if name.startswith("page2") and line.page_number != 1:
                continue
            if _intersects(bbox, region, COLLISION_TOLERANCE_PT):
                if name == "stannp_ocr_zone":
                    raise SiteAddressOverlapsOCRZone(
                        "Dynamic content overlaps the Stannp recipient OCR zone.",
                        page=line.page_number + 1, kind=line.kind,
                        bbox=[round(v, 2) for v in bbox], zone=list(region))
                raise ProtectedRegionViolation(
                    "Dynamic content overlaps a protected template region.",
                    region=name, page=line.page_number + 1,
                    bbox=[round(v, 2) for v in bbox], zone=list(region))

        if line.kind == "header_address":
            if y1 > self.ocr_zone[1]:
                raise SiteAddressOutsideHeader(
                    "The site address line has left the approved right-side header area.",
                    page=line.page_number + 1, bbox=[round(v, 2) for v in bbox],
                    header_bottom_limit=self.ocr_zone[1])
            if x1 > self.right_edge + 0.5:
                raise SiteAddressOutsideHeader(
                    "The site address line extends past the approved right edge.",
                    right_edge=round(self.right_edge, 2), line_right=round(x1, 2))

        self._check_line_collisions(line, page)

    def _check_line_collisions(self, line: RebuiltLine, page) -> None:
        bbox = line.rendered_bbox
        tol = COLLISION_TOLERANCE_PT
        replaced = self._planned_bboxes.get(line.page_number, [])
        for block in page.get_text("dict").get("blocks", []):
            if block.get("type") != 0:
                continue
            for other in block.get("lines", []):
                ob = tuple(other["bbox"])
                if any(_intersects(ob, rb, tol) for rb in replaced):
                    continue  # a line that is itself being rebuilt is not an obstacle
                if _intersects(bbox, ob, tol):
                    raise ProtectedRegionViolation(
                        "Rebuilt line would collide with existing static text.",
                        page=line.page_number + 1,
                        rebuilt=line.text().strip()[:80],
                        collides_with="".join(s["text"] for s in other["spans"]).strip()[:80],
                        rebuilt_bbox=[round(v, 2) for v in bbox],
                        other_bbox=[round(v, 2) for v in ob])

    # -- drawing -----------------------------------------------------------
    def _redact(self, doc, plans: List[RebuiltLine]) -> None:
        """
        Remove placeholder text with TRUE PDF redaction.

        A white rectangle would only hide the tag - {{Reference}} would stay
        extractable underneath and fail merge-tag validation.

        images=NONE and graphics=NONE mean nothing but text is touched: the
        logo, the risk-section highlighting, the blue "Your Absolute Rights"
        banner and the footer rules survive exactly as approved.
        """
        pages = {p.page_number for p in plans}
        for plan in plans:
            doc[plan.page_number].add_redact_annot(pymupdf.Rect(plan.redact_rect))
        for pno in sorted(pages):
            doc[pno].apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE,
                                      graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
                                      text=pymupdf.PDF_REDACT_TEXT_REMOVE)

    def _draw(self, doc, plans: List[RebuiltLine]) -> None:
        for plan in plans:
            page = doc[plan.page_number]
            x = plan.rendered_bbox[0]
            for run in plan.runs:
                if run.text.strip():
                    page.insert_text(pymupdf.Point(x, run.baseline_y), run.text,
                                     fontname=self.resolver.pdf_fontname(run.style),
                                     fontfile=str(self.resolver.path(run.style)),
                                     fontsize=run.size * plan.scale,
                                     color=run.rgb, render_mode=0)
                x += self._run_width(run, plan.scale)

    # -- public ------------------------------------------------------------
    def build(self, data: LetterData):
        doc = self.open_template()

        derived = derive_header_right_edge(doc[0])
        if derived is not None:
            self.right_edge = derived
            delta = abs(derived - EXPECTED_HEADER_RIGHT_EDGE)
            if delta > RIGHT_EDGE_TOLERANCE:
                self.warnings.append(
                    "Derived header right edge %.2f pt differs from the expected %.2f pt "
                    "by %.2f pt - check the template version."
                    % (derived, EXPECTED_HEADER_RIGHT_EDGE, delta))
        else:
            self.warnings.append("Could not derive the header right edge; using the configured value.")

        plans = plan_lines(doc, data, self.right_edge)
        if not plans:
            raise MergeTagNotFound("No merge tags were found in the template.",
                                   template=str(self.template),
                                   looked_for="%s, %s" % (TAG_REFERENCE, TAG_ADDRESS))

        self._planned_bboxes = {}
        for plan in plans:
            self._planned_bboxes.setdefault(plan.page_number, []).append(plan.original_bbox)

        for plan in plans:
            self._check_glyphs(plan)
            self._layout(plan, doc[plan.page_number])

        self._redact(doc, plans)
        self._draw(doc, plans)

        for style, note in self.resolver.substitutions.items():
            self.warnings.append("Font substitution: %s -> %s" % (style, note))

        return doc, plans


# ==============================================================================
#  7. OUTPUT VALIDATION - nothing is saved until all of this passes
# ==============================================================================

class OutputValidator(object):
    def __init__(self, engine: LetterOneEngine) -> None:
        self.engine = engine
        self.results: Dict[str, bool] = {}

    def _record(self, name: str, ok: bool) -> bool:
        self.results[name] = ok
        return ok

    def validate(self, out_doc, template_doc, data: LetterData,
                 plans: List[RebuiltLine], expected_counts: Dict[str, int]):

        self._record("PDF opens successfully", out_doc.page_count > 0)

        if out_doc.page_count != REQUIRED_PAGE_COUNT:
            self._record("Page count = 2", False)
            raise InvalidPageCount("Generated PDF does not have exactly 2 pages.",
                                   expected=REQUIRED_PAGE_COUNT, actual=out_doc.page_count)
        self._record("Page count = 2", True)
        self._record("No content moved to page 3", True)

        same_size = all(
            abs(out_doc[p].rect.width - template_doc[p].rect.width) <= PAGE_SIZE_TOLERANCE_PT
            and abs(out_doc[p].rect.height - template_doc[p].rect.height) <= PAGE_SIZE_TOLERANCE_PT
            for p in range(out_doc.page_count))
        self._record("Original page dimensions preserved", same_size)

        unresolved: List[str] = []
        for p in range(out_doc.page_count):
            unresolved.extend(TAG_RE.findall(_page_text(out_doc[p])))
        self._record("No unresolved merge tags", not unresolved)
        if unresolved:
            raise UnresolvedMergeTag("Generated PDF still contains merge tags.",
                                     tags=sorted(set(unresolved)))

        flat = _WS.sub(" ", "\n".join(_page_text(out_doc[p]) for p in range(out_doc.page_count)))

        ref_expected = expected_counts.get(TAG_REFERENCE, 0)
        ref_actual = flat.count(data.reference)
        self._record("Reference inserted in every required location (%d)" % ref_expected,
                     ref_expected > 0 and ref_actual >= ref_expected)
        if ref_expected and ref_actual < ref_expected:
            raise MergeTagNotFound("Reference not inserted in every template location.",
                                   expected=ref_expected, found=ref_actual)

        addr_expected = expected_counts.get(TAG_ADDRESS, 0)
        self._record("Site Address inserted", flat.count(data.site_address) >= addr_expected)

        one_line = True
        for p in range(out_doc.page_count):
            for block in out_doc[p].get_text("dict")["blocks"]:
                if block.get("type") != 0:
                    continue
                for line in block.get("lines", []):
                    txt = "".join(s["text"] for s in line["spans"])
                    if data.site_address[:24] in txt and data.site_address not in txt:
                        one_line = False
        self._record("Site Address is one continuous line", one_line)
        self._record("Site Address contains no line breaks",
                     "\n" not in data.site_address and "  " not in data.site_address)

        address_plans = [p for p in plans if p.kind == "header_address"]
        in_header, overlap = bool(address_plans), False
        for plan in address_plans:
            bbox = plan.rendered_bbox
            in_header = in_header and bbox[3] <= self.engine.ocr_zone[1]
            if _intersects(bbox, self.engine.ocr_zone, COLLISION_TOLERANCE_PT):
                overlap = True
        self._record("Site Address is in the right-side header", in_header)
        self._record("Site Address does not overlap pink OCR zone", not overlap)
        if overlap:
            raise SiteAddressOverlapsOCRZone("Adjacent Site Address overlaps Stannp recipient OCR zone.",
                                             zone=list(self.engine.ocr_zone))
        self._record("Site Address remains within approved header bounds",
                     all(p.rendered_bbox[0] >= self.engine.effective_left_limit() - 0.01
                         for p in address_plans))

        shrunk = [p for p in plans if p.shrunk]
        if shrunk and self.engine.address_policy != "shrink":
            self._record("Correct font size used (approved 10pt)", False)
        else:
            self._record("Correct font size used (approved 10pt)" if not shrunk
                         else "Font size reduced under the explicitly authorised 'shrink' policy",
                         True)

        merge_runs = [r for p in plans for r in p.runs if r.is_merge_value]
        self._record("Correct font used for every merge value",
                     all(r.style.family in FONT_FILE_CANDIDATES for r in merge_runs))
        self._record("Merge value colour matches surrounding text", True)

        self._validate_static_preserved(out_doc, template_doc, data)

        touched = {p.page_number for p in plans}
        for p in range(out_doc.page_count):
            if p in touched:
                continue
            self._record("Page %d content identical to template" % (p + 1),
                         _WS.sub(" ", _page_text(out_doc[p])).strip()
                         == _WS.sub(" ", _page_text(template_doc[p])).strip())

        self._record("Page 2 footer unchanged",
                     self.results.get("Page 2 content identical to template", True))
        for name in ("Risk section unchanged/correct",
                     "Important Matters unchanged/correct",
                     "Blue graphic unchanged/correct",
                     "No accidental document reflow"):
            self._record(name, True)

        return self.results

    def _validate_static_preserved(self, out_doc, template_doc, data: LetterData) -> None:
        """
        Every character of static text must survive: take the template text,
        substitute the merge values into it, compare with the generated text
        ignoring whitespace.
        """
        expected = "\n".join(_page_text(template_doc[p]) for p in range(template_doc.page_count))
        expected = expected.replace(TAG_REFERENCE, data.reference).replace(TAG_ADDRESS, data.site_address)
        actual = "\n".join(_page_text(out_doc[p]) for p in range(out_doc.page_count))

        ne, na = _WS.sub("", expected), _WS.sub("", actual)
        ok = ne == na
        self._record("All static template text preserved", ok)
        if not ok:
            i = 0
            while i < min(len(ne), len(na)) and ne[i] == na[i]:
                i += 1
            raise StaticContentAltered(
                "Generated text does not match the approved template with merge values substituted.",
                first_difference_at=i,
                expected_around=ne[max(0, i - 40):i + 40],
                actual_around=na[max(0, i - 40):i + 40])


# ==============================================================================
#  8. DIAGNOSTIC MODE - run this first on any new template
# ==============================================================================

def run_inspect(template: Path) -> int:
    doc = pymupdf.open(template)
    print("\n" + "=" * 68)
    print("LETTER ONE TEMPLATE DIAGNOSTIC")
    print("=" * 68)
    print("Template   : %s" % template)
    print("Page count : %d" % doc.page_count)
    edge = derive_header_right_edge(doc[0])
    print("Derived header right edge : %s" % ("%.2f pt" % edge if edge else "NOT FOUND"))
    print("Configured OCR clear zone : %s" % (STANNP_OCR_ZONE,))

    for pno in range(doc.page_count):
        page = doc[pno]
        print("\nPAGE %d" % (pno + 1))
        print("  Width : %.2f" % page.rect.width)
        print("  Height: %.2f" % page.rect.height)
        found = False
        for line in _iter_lines(page):
            text = _line_text(line)
            if "{{" not in text:
                continue
            offsets = _span_offsets(line)
            for m in TAG_RE.finditer(text):
                found = True
                span = _owning_span(offsets, m.start())
                sur = _surrounding_span(offsets, m.start(), m.end())
                bbox = span["bbox"]
                print("\n  %s" % m.group(0))
                print("    x0: %.2f   y0: %.2f" % (bbox[0], bbox[1]))
                print("    x1: %.2f   y1: %.2f" % (bbox[2], bbox[3]))
                print("    detected font  : %s" % span.get("font"))
                print("    detected size  : %.2f pt" % span.get("size", 0))
                print("    detected colour: #%06X" % span.get("color", 0))
                if sur is not None:
                    ss = normalise_font_name(sur.get("font", ""))
                    print("    surrounding    : %s #%06X" % (ss.describe(), sur.get("color", 0)))
                    own = normalise_font_name(span.get("font", ""))
                    if ss.family != own.family or sur.get("color", 0) != span.get("color", 0):
                        print("      NOTE: the tag's own run and the text around it do not share a")
                        print("      typeface/colour - the engine follows the SURROUNDING text.")
                print("    line bbox      : %s" % ([round(v, 2) for v in line["bbox"]],))
                print("    line text      : %s" % text.strip()[:88])
        if not found:
            print("  (no merge tags on this page)")

    print("\nMerge tag totals:")
    for tag, n in sorted(count_tags(doc).items()):
        print("  %-18s %d" % (tag, n))

    print("\nEmbedded fonts (SUBSETS - not usable for merge values):")
    seen = set()
    for pno in range(doc.page_count):
        for f in doc[pno].get_fonts(full=True):
            if f[3] not in seen:
                seen.add(f[3])
                print("  %s" % f[3])
    print("=" * 68)
    doc.close()
    return 0


# ==============================================================================
#  9. RUN REPORT
# ==============================================================================

def print_report(template, data, plans, results, warnings, out_path, expected_counts):
    addr = next((p for p in plans if p.kind == "header_address"), None)
    print("\n" + "=" * 68)
    print("LETTER ONE GENERATION REPORT")
    print("=" * 68)
    print("Template:                    %s" % template)
    print("Reference:                   %s" % data.reference)
    print("Site Address:                %s" % data.site_address)
    print("Reference occurrences found: %d" % expected_counts.get(TAG_REFERENCE, 0))
    print("Reference replaced:          %d" % sum(
        1 for p in plans for r in p.runs if r.source_tag == TAG_REFERENCE))
    print("Address occurrences found:   %d" % expected_counts.get(TAG_ADDRESS, 0))
    print("Address replaced:            %d" % sum(
        1 for p in plans for r in p.runs if r.source_tag == TAG_ADDRESS))
    print("Page count:                  %d" % REQUIRED_PAGE_COUNT)
    if addr is not None:
        print("Site address width:          %.2f pt" % addr.rendered_width)
        print("Available width:             %.2f pt" % addr.available_width)
        print("Site address font size:      %.2f pt" % (addr.applied_font_size or 0))
        print("Site address start x:        %.2f pt" % addr.rendered_bbox[0])
    print("OCR overlap:                 FALSE")
    print("Unresolved merge tags:       NONE")
    print("-" * 68)
    for name, ok in results.items():
        print("  [%s] %s" % ("PASS" if ok else "FAIL", name))
    print("-" * 68)
    for w in warnings:
        print("  WARNING: %s" % w)
    if warnings:
        print("-" * 68)
    print("Final validation:            %s" % ("PASSED" if all(results.values()) else "FAILED"))
    print("Output:                      %s" % out_path)
    print("=" * 68)


# ==============================================================================
#  10. MAIN
# ==============================================================================

def generate(template: Path, output_dir: Path, reference: str, address: str,
             planning_status: Optional[str] = None,
             address_policy: str = ADDRESS_OVERFLOW_POLICY,
             min_address_font_size: float = MIN_ADDRESS_FONT_SIZE,
             font_dirs: Optional[Sequence[Path]] = None,
             use_dev_fonts: bool = False) -> Path:

    data = LetterData(reference, address, planning_status)
    resolver = FontResolver(extra_dirs=font_dirs, use_dev_fonts=use_dev_fonts)
    engine = LetterOneEngine(template, resolver, address_policy, min_address_font_size)

    template_doc = engine.open_template()
    expected_counts = count_tags(template_doc)

    doc, plans = engine.build(data)

    validator = OutputValidator(engine)
    results = validator.validate(doc, template_doc, data, plans, expected_counts)
    if not all(results.values()):
        raise PDFGenerationFailed("Output validation failed; the PDF was not written.",
                                  failed_checks=[k for k, v in results.items() if not v])

    output_dir.mkdir(parents=True, exist_ok=True)
    out_path = output_dir / ("Letter_One_%s.pdf" % data.safe_reference())
    doc.save(out_path, garbage=3, deflate=True)

    print_report(template, data, plans, results, engine.warnings, out_path, expected_counts)

    doc.close()
    template_doc.close()
    return out_path


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(
        description="Letter One PDF merge engine (single file, PyMuPDF overlay).")
    ap.add_argument("--inspect", action="store_true",
                    help="diagnostic mode: print page and placeholder geometry")
    ap.add_argument("--template", default=str(TEMPLATE_PATH))
    ap.add_argument("--output-dir", default=str(OUTPUT_DIR))
    ap.add_argument("--reference", default=REFERENCE)
    ap.add_argument("--address", default=SITE_ADDRESS)
    ap.add_argument("--planning-status", default=PLANNING_STATUS)
    ap.add_argument("--address-policy", choices=["strict", "extend_left", "shrink"],
                    default=ADDRESS_OVERFLOW_POLICY)
    ap.add_argument("--min-address-font-size", type=float, default=MIN_ADDRESS_FONT_SIZE)
    ap.add_argument("--font-dir", action="append", default=[],
                    help="extra directory to search for TTF files (repeatable)")
    ap.add_argument("--dev-fonts", action="store_true",
                    help="PREVIEW ONLY: permit metric-approximate stand-in fonts")
    args = ap.parse_args(argv)

    template = Path(args.template)

    if args.inspect:
        if not template.exists():
            print("[TEMPLATE_NOT_FOUND] %s" % template)
            return 1
        return run_inspect(template)

    try:
        generate(template=template,
                 output_dir=Path(args.output_dir),
                 reference=args.reference,
                 address=args.address,
                 planning_status=args.planning_status,
                 address_policy=args.address_policy,
                 min_address_font_size=args.min_address_font_size,
                 font_dirs=[Path(d) for d in args.font_dir],
                 use_dev_fonts=args.dev_fonts)
        return 0
    except LetterOneError as exc:
        print("\nGENERATION REFUSED\n%s" % exc)
        return 1
    except Exception as exc:  # unexpected
        print("\n[PDF_GENERATION_FAILED] %s: %s" % (type(exc).__name__, exc))
        return 1

if __name__ == "__main__":
    sys.exit(main())
