"""
Fetch static site generator — Paw Sweet Paw pilot.

Reads the Markdown source of truth (company-defaults.md, brand-knowledge/,
location-knowledge/) and renders self-contained HTML pages for publishing.
This is a one-way generator: the HTML is always regenerated from Markdown,
never hand-edited. Re-run after any Markdown change (manual edit or an
approved change from email_processor.py) and redeploy.

Schema follows Taylor's dashboard spec (Inputs/TAYLOR_DASHBOARD_VIEW.md):
six tabs per location — Center Information, Boarding, Daycare, Training,
Grooming, Policies & Procedures — each broken into a fixed set of tiles.
Every location page (backend knowledgebase view) and dashboard page (at-a-
glance view) is built from the SAME schema (FULL_SCHEMA below), so a topic
lives at one heading in the location .md file and is rendered consistently
in both places. A missing topic is a visible "not on file" gap rather than
a silently missing section. Gaps are also collected into gaps.json so they
can feed a future missing-info-bank workflow.

Vaccines, cameras, and add-ons are each documented ONCE per location (under
Boarding) and rendered into both the Boarding and Daycare tab/section —
single source of truth, shown in multiple places, per the "one source,
multiple tabs" pattern (rather than duplicating the same facts twice in the
source file and risking them drifting out of sync).

Center Information is the one tab that mixes two source types: Contact
Info / Org Chart / Business Links / Hours pull from the ops CSV (system,
kept live outside the Markdown chain) on the DASHBOARD page, with
Discounts and Tours pulled from the location .md file (narrative facts the
CSV doesn't track). The backend knowledgebase page sources Center
Information purely from the .md file, consistent with how every other
topic on that page works.
"""

import os
import re
import csv
import json
import shutil
import markdown
import openpyxl
from PIL import Image, ImageChops

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SITE_ROOT = os.path.join(BASE_DIR, "fetch-app", "site")
PSP_DIR = os.path.join(SITE_ROOT, "psp")

# Placeholder — update once the real Netlify/custom domain is chosen (see plan, step 2).
SITE_BASE_URL = "https://fetch-paw-sweet-paw.netlify.app"

# Company-wide ops CSV (center directory) — a separate pipeline from the
# Markdown knowledge base, not part of the location/brand/standard
# inheritance chain. Manually exported snapshot for now (filename carries
# the pull date); swapping this loader for a live Google Sheets pull is a
# planned upgrade, not built yet — see fetch-context.md next steps.
CENTER_INFO_CSV = os.path.join(BASE_DIR, "Inputs", "Center_info_8.10.csv")
CENTER_INFO_XLSX = os.path.join(BASE_DIR, "Inputs", "Center_Info_8.10.xlsx")

# POS-branded "visit this center's booking site" button on the dashboard —
# logo shown depends on which POS system the center runs on. MoeGo is the
# default; Gingr is the only exception today (The Farm), per Tucker.
POS_LOGO_ASSETS = {
    "gingr": {"src": os.path.join(BASE_DIR, "Inputs", "gingr logo.png"), "asset_name": "gingr-logo.png", "alt": "Gingr"},
    "moego": {"src": os.path.join(BASE_DIR, "Inputs", "moego logo.jpg"), "asset_name": "moego-logo.jpg", "alt": "MoeGo"},
}


def _trim_whitespace(img, pad=30):
    """MoeGo's source logo has a lot of empty margin baked into the file,
    which makes the wordmark tiny once scaled to button height. Crop to the
    actual artwork (plus a small pad) so it reads clearly at small sizes."""
    bg = Image.new(img.mode, img.size, (255, 255, 255) if img.mode != "RGBA" else (255, 255, 255, 255))
    bbox = ImageChops.difference(img.convert("RGB"), bg.convert("RGB")).getbbox()
    if not bbox:
        return img
    w, h = img.size
    l, t, r, b = bbox
    return img.crop((max(0, l - pad), max(0, t - pad), min(w, r + pad), min(h, b + pad)))


def copy_pos_logo_assets():
    assets_dir = os.path.join(SITE_ROOT, "assets")
    os.makedirs(assets_dir, exist_ok=True)
    for key, logo in POS_LOGO_ASSETS.items():
        dest = os.path.join(assets_dir, logo["asset_name"])
        if key == "moego":
            _trim_whitespace(Image.open(logo["src"])).convert("RGB").save(dest, quality=92)
        else:
            shutil.copyfile(logo["src"], dest)


def build_pos_booking_button(row):
    """Button showing the correct POS logo (Gingr vs MoeGo) linking to this
    center's internal POS Direct Login (staff access) — not the customer
    booking portal, which lives in the Business Links tile instead."""
    if not row:
        return ""
    book_url = (row.get("POS_Direct_Login_URL") or "").strip()
    if not book_url.startswith("http"):
        return ""
    pos = (row.get("POS System") or "").strip().lower()
    logo = POS_LOGO_ASSETS["gingr"] if "gingr" in pos else POS_LOGO_ASSETS["moego"]
    return (
        f'<a class="pos-booking-btn" href="{book_url}" target="_blank" rel="noopener">'
        f'<img src="../assets/{logo["asset_name"]}" alt="{logo["alt"]} logo">'
        f"Visit this center's booking site</a>"
    )


def load_center_info():
    with open(CENTER_INFO_CSV, encoding="utf-8-sig", newline="") as f:
        info = {row["Center_ID"]: row for row in csv.DictReader(f)}

    # CSV export flattens hyperlinks to display text ("MoeGo Link") — pull the
    # real POS Direct Login target from column N of the source workbook,
    # since that's the one field that's actually useful as a link.
    wb = openpyxl.load_workbook(CENTER_INFO_XLSX)
    ws = wb.active
    for row in ws.iter_rows(min_row=2):
        center_id = row[0].value
        if center_id is None:
            continue
        center_id = str(int(center_id)) if isinstance(center_id, float) else str(center_id)
        cell = row[13]  # column N: POS_Direct_Login
        if center_id in info and cell.hyperlink:
            info[center_id]["POS_Direct_Login_URL"] = cell.hyperlink.target

    return info


MD = markdown.Markdown(extensions=["tables", "fenced_code"])

LOCATIONS = [
    {"file": "246-paw-sweet-paw-jamboree.md", "slug": "jamboree", "name": "Paw Sweet Paw Jamboree", "center_id": "246"},
    {"file": "247-paw-sweet-paw-spectrum.md", "slug": "spectrum", "name": "Paw Sweet Paw Spectrum", "center_id": "247"},
]

# ── Taylor's dashboard schema (Inputs/TAYLOR_DASHBOARD_VIEW.md) ──────────
# One entry per tab: (tab_key, tab_title, [(tile_key, tile_title, location_h3_heading), ...]).
# "COMPOSITE_VACCINE_TOGETHER" is a sentinel handled specially in resolve_tile_content —
# it combines the location file's single "Vaccine Requirements" and "Boarding
# Together Rules" h3 blocks into one tile, reused under both Boarding and
# Daycare with a different tile title in each ("...Room Sharing..." vs
# "...Spay/Neuter...") since the same underlying facts answer both framings.
FULL_SCHEMA = [
    ("center-info", "Center Information", [
        ("contact-info", "Contact Info", "Contact Info"),
        ("org-chart", "Organizational Chart", "Organizational Chart"),
        ("discounts", "Discounts Offered", "Discounts Offered"),
        ("hours", "Hours of Operation", "Hours of Operation"),
        ("business-links", "Business Links", "Business Links"),
        ("tours", "Tour Information", "Tour Information"),
    ]),
    ("boarding", "Boarding", [
        ("dog-accommodation", "Dog Accommodation Options", "Dog Accommodation Options"),
        ("cat-accommodation", "Cat Accommodation Options", "Cat Accommodation Options"),
        ("pocket-pet-accommodation", "Pocket Pet Accommodation Options", "Pocket Pet Accommodation Options"),
        ("vaccine-room-sharing", "Vaccine & Room Sharing Requirements", "COMPOSITE_VACCINE_TOGETHER"),
        ("boarding-addons", "Add-on/Enrichment Options", "Add-On Services"),
        ("boarding-cameras", "Cameras", "Camera Access"),
    ]),
    ("daycare", "Daycare", [
        ("daycare-pricing", "Daycare Pricing + Package Options", "Daycare Pricing & Packages"),
        ("group-play", "Group Play Details", "Group Play Details"),
        ("evaluation", "Evaluation Details", "Evaluation Details"),
        ("vaccine-spay-neuter", "Vaccine & Spay/Neuter Requirements", "COMPOSITE_VACCINE_TOGETHER"),
        ("daycare-addons", "Add-on/Enrichment Options", "Add-On Services"),
        ("daycare-cameras", "Cameras", "Camera Access"),
    ]),
    ("training", "Training", [
        ("center-trainer", "Center Trainer", "Center Trainer"),
        ("board-train", "Board + Train", "Board + Train"),
        ("single-session", "Single Session Training", "Single Session Training"),
        ("training-interest-form", "Dog Training Interest Form", "Dog Training Interest Form"),
    ]),
    ("grooming", "Grooming", [
        ("groomers", "Groomers", "Groomers"),
        ("bathers", "Bathers", "Bathers"),
        ("grooming-pricing", "Grooming & Bathing Pricing", "Grooming & Bathing Pricing"),
        ("grooming-addons", "Grooming Add-On/Package Options", "Grooming Add-On/Package Options"),
        ("grooming-hours", "Hours of Operation", "Grooming Hours of Operation"),
        ("grooming-misc", "Misc.", "Misc."),
    ]),
    ("policies", "Policies & Procedures", [
        ("bfpc-policy-docs", "BFPC Policy Docs", "BFPC Policy Docs"),
    ]),
]

# (cols, rows) on the desktop 3-column dashboard grid. A tile keeps this
# footprint at every breakpoint (clamped to however many columns actually
# exist at that width).
TILE_SPANS = {
    # Center Information
    "contact-info": (1, 1), "org-chart": (1, 1), "discounts": (1, 1),
    "hours": (2, 1), "business-links": (1, 1), "tours": (3, 1),
    # Boarding
    "dog-accommodation": (1, 1), "cat-accommodation": (1, 1), "pocket-pet-accommodation": (1, 1),
    "vaccine-room-sharing": (2, 2), "boarding-addons": (1, 1), "boarding-cameras": (1, 1),
    # Daycare
    "daycare-pricing": (2, 1), "group-play": (1, 1),
    "evaluation": (1, 2), "vaccine-spay-neuter": (1, 2),
    "daycare-addons": (1, 1), "daycare-cameras": (1, 1),
    # Training
    "center-trainer": (1, 1), "board-train": (1, 1), "single-session": (1, 1), "training-interest-form": (3, 1),
    # Grooming
    "groomers": (1, 1), "bathers": (1, 1), "grooming-hours": (1, 1),
    "grooming-pricing": (2, 2), "grooming-addons": (1, 1), "grooming-misc": (1, 1),
    # Policies & Procedures
    "bfpc-policy-docs": (3, 1),
}

PAGE_TEMPLATE = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title} — Paw Sweet Paw Knowledgebase</title>
<meta name="description" content="{description}">
<meta name="robots" content="index, follow">
<link rel="canonical" href="{canonical}">
{jsonld}
<style>
  :root {{ color-scheme: light dark; }}
  body {{ font-family: -apple-system, Segoe UI, Roboto, sans-serif; max-width: 820px; margin: 0 auto; padding: 32px 20px 80px; line-height: 1.55; color: #1c1c1e; background: #fff; }}
  @media (prefers-color-scheme: dark) {{ body {{ color: #e8e8ea; background: #16161a; }} }}
  h1 {{ font-size: 1.6rem; margin-bottom: 4px; }}
  h2 {{ font-size: 1.2rem; margin-top: 0; border-bottom: 1px solid #ddd; padding-bottom: 4px; scroll-margin-top: 12px; }}
  @media (prefers-color-scheme: dark) {{ h2 {{ border-color: #333; }} }}
  h3 {{ font-size: 1.02rem; margin-top: 1.4em; }}
  table {{ border-collapse: collapse; width: 100%; margin: 0.8em 0; font-size: 0.95rem; }}
  th, td {{ border: 1px solid #ddd; padding: 6px 10px; text-align: left; }}
  @media (prefers-color-scheme: dark) {{ th, td {{ border-color: #333; }} }}
  th {{ background: #f5f5f7; }}
  @media (prefers-color-scheme: dark) {{ th {{ background: #222226; }} }}
  .nav {{ font-size: 0.9rem; margin-bottom: 18px; }}
  .nav a {{ color: #0a6cff; text-decoration: none; margin-right: 14px; }}
  .toc {{ font-size: 0.88rem; background: #f7f7f9; border: 1px solid #e4e4e8; border-radius: 8px; padding: 14px 16px; margin: 18px 0 28px; }}
  @media (prefers-color-scheme: dark) {{ .toc {{ background: #1e1e22; border-color: #2c2c31; }} }}
  .toc strong {{ display: block; margin-bottom: 10px; font-size: 0.78rem; text-transform: uppercase; letter-spacing: 0.04em; color: #888; }}
  .toc-links {{ display: flex; flex-wrap: wrap; gap: 8px 8px; }}
  .toc a, .toc-links a {{ display: inline-block; color: #24478f; background: #e8eefc; text-decoration: none; white-space: nowrap; padding: 5px 12px; border-radius: 999px; font-size: 0.82rem; border: 1px solid transparent; transition: background 0.15s; }}
  .toc a:hover, .toc-links a:hover {{ background: #d6e0fa; }}
  @media (prefers-color-scheme: dark) {{
    .toc a, .toc-links a {{ color: #8fb0f7; background: #16213f; }}
    .toc a:hover, .toc-links a:hover {{ background: #1e2d54; }}
  }}
  .toc-links a.gap-link {{ color: #9c2b2b; background: #fbe7e7; }}
  @media (prefers-color-scheme: dark) {{ .toc-links a.gap-link {{ color: #f0a3a3; background: #3a1717; }} }}
  .source-banner {{ display: inline-block; font-size: 0.72rem; font-weight: 600; letter-spacing: 0.03em; text-transform: uppercase; padding: 2px 8px; border-radius: 4px; margin: 0 0 6px; }}
  .source-location {{ background: #e6f4ea; color: #1a7431; }}
  .source-brand {{ background: #e8eefc; color: #24478f; }}
  .source-default {{ background: #f4eee3; color: #7a5b1e; }}
  .source-gap {{ background: #fbe7e7; color: #9c2b2b; }}
  @media (prefers-color-scheme: dark) {{
    .source-location {{ background: #10331b; color: #6fd98c; }}
    .source-brand {{ background: #16213f; color: #8fb0f7; }}
    .source-default {{ background: #332a15; color: #e0bb6a; }}
    .source-gap {{ background: #3a1717; color: #f0a3a3; }}
  }}
  section {{ margin-top: 2.4em; padding-top: 1.3em; border-top: 1px solid #e4e4e8; scroll-margin-top: 12px; }}
  @media (prefers-color-scheme: dark) {{ section {{ border-color: #2c2c31; }} }}
  section:first-of-type {{ margin-top: 0; padding-top: 0; border-top: none; }}
  .group-header {{ font-size: 1.4rem; margin: 2.6em 0 0.3em; scroll-margin-top: 12px; }}
  .group-header:first-of-type {{ margin-top: 1.6em; }}
  .gap-note {{ font-style: italic; color: #9c2b2b; }}
  @media (prefers-color-scheme: dark) {{ .gap-note {{ color: #f0a3a3; }} }}
  .reference-note {{ color: #555; }}
  @media (prefers-color-scheme: dark) {{ .reference-note {{ color: #aaa; }} }}
  blockquote {{ border-left: 3px solid #0a6cff; margin: 0.8em 0; padding: 2px 14px; color: #444; }}
  @media (prefers-color-scheme: dark) {{ blockquote {{ color: #bbb; }} }}
  footer {{ margin-top: 3em; font-size: 0.8rem; color: #888; border-top: 1px solid #ddd; padding-top: 12px; }}
  @media (prefers-color-scheme: dark) {{ footer {{ border-color: #333; }} }}
</style>
</head>
<body>
{nav}
{body}
<footer>
  Generated from the Fetch knowledge base (company defaults → Paw Sweet Paw brand → location). Every fact is tagged by source level — Location, Brand, Best Friends Standard, or a flagged Gap if no level has it on file. Company-wide policy lives once on the Best Friends Standard page and is linked from here, not duplicated. This page is generated automatically and should not be edited directly; corrections should go through the Fetch update process.
</footer>
</body>
</html>
"""


def build_nav(context):
    """context: 'psp' for pages inside /psp/, 'root' for the shared standards page."""
    if context == "psp":
        links = [
            ("index.html", "Paw Sweet Paw — Overview"),
            ("jamboree.html", "Jamboree"),
            ("spectrum.html", "Spectrum"),
            ("../best-friends-standard.html", "Best Friends Standard"),
        ]
    else:
        links = [
            ("psp/index.html", "Paw Sweet Paw — Overview"),
            ("psp/jamboree.html", "Jamboree"),
            ("psp/spectrum.html", "Spectrum"),
            ("best-friends-standard.html", "Best Friends Standard"),
        ]
    inner = "".join(f'<a href="{href}">{label}</a>' for href, label in links)
    return f'<div class="nav">{inner}</div>'


def ensure_blank_line_before_tables(text):
    """python-markdown only recognizes a table if it starts a new block."""
    lines = text.split("\n")
    out = []
    for line in lines:
        if line.startswith("|") and out and out[-1].strip() and not out[-1].startswith("|"):
            out.append("")
        out.append(line)
    return "\n".join(out)


def convert(md_text):
    return MD.reset().convert(ensure_blank_line_before_tables(md_text))


BANNER_CLASS = {"location": "source-location", "brand": "source-brand", "default": "source-default", "gap": "source-gap", "system": "source-system"}
BANNER_TEXT = {
    "location": "Location-specific",
    "brand": "Paw Sweet Paw brand standard",
    "default": "Best Friends standard",
    "gap": "Not on file",
    "system": "Center info system (CSV)",
}


def render_section(anchor_key, title, source_label, html_or_gap_text):
    banner = f'<div class="source-banner {BANNER_CLASS[source_label]}">{BANNER_TEXT[source_label]}</div>'
    # A <div> wrapper (not <p>) — gap content is already-converted markdown
    # that may itself contain <p> tags, and nesting <p> inside <p> is invalid HTML.
    body = f'<div class="gap-note">{html_or_gap_text}</div>' if source_label == "gap" else html_or_gap_text
    return f'<section id="{anchor_key}">{banner}<h2>{title}</h2>{body}</section>'


def strip_frontmatter(text):
    return re.sub(r"^---\n.*?\n---\n", "", text, flags=re.DOTALL)


def read(path):
    with open(path, encoding="utf-8") as f:
        return strip_frontmatter(f.read())


def extract_h2_block(full_text, heading):
    """Pull one ## section (through the next ## or end of file) by heading text."""
    pattern = rf"^## {re.escape(heading)}\n(.*?)(?=^## |\Z)"
    m = re.search(pattern, full_text, flags=re.MULTILINE | re.DOTALL)
    return m.group(1).strip() if m else None


def extract_h3_block(full_text, heading):
    pattern = rf"^### {re.escape(heading)}\n(.*?)(?=^#{{2,3}} |\Z)"
    m = re.search(pattern, full_text, flags=re.MULTILINE | re.DOTALL)
    return m.group(1).strip() if m else None


def extract_field(block, field_label):
    m = re.search(rf"\*\*{re.escape(field_label)}:\*\*\s*(.+)", block or "")
    return m.group(1).strip() if m else None


def build_toc(items):
    links = "".join(f'<a href="#{key}">{title}</a>' for key, title, *_ in items)
    return f'<nav class="toc"><strong>On this page</strong><div class="toc-links">{links}</div></nav>'


GAP_PREFIXES = ("_Not on file", "_Not offered", "_Not documented")


def is_gap_text(block):
    return bool(block) and block.strip().startswith(GAP_PREFIXES)


def resolve_tile_content(loc_text, heading):
    """Returns (html, source_label) for one tile, sourced purely from the
    location .md file. source_label is 'gap' if the heading is missing
    entirely OR the block on file is itself a "not on file" note — either
    way it renders with the same red gap styling and gets logged."""
    if heading == "COMPOSITE_VACCINE_TOGETHER":
        vaccines = extract_h3_block(loc_text, "Vaccine Requirements")
        together = extract_h3_block(loc_text, "Boarding Together Rules")
        if not vaccines and not together:
            return None, "gap"
        parts = []
        gap = False
        if vaccines:
            parts.append(f"<h3>Vaccines</h3>{convert(vaccines)}")
            gap = gap or is_gap_text(vaccines)
        else:
            gap = True
        if together:
            parts.append(f"<h3>Room Sharing / Group Play</h3>{convert(together)}")
            gap = gap or is_gap_text(together)
        else:
            gap = True
        return "".join(parts), ("gap" if gap else "location")

    block = extract_h3_block(loc_text, heading)
    if block is None:
        return None, "gap"
    return convert(block), ("gap" if is_gap_text(block) else "location")


def build_tab_sections(loc_text, tab_tiles, gaps_log, loc_slug):
    """Backend knowledgebase page: one <section> per tile, all sourced from
    the location .md file."""
    out = []
    for key, title, heading in tab_tiles:
        html, source = resolve_tile_content(loc_text, heading)
        if html is None:
            gaps_log.append({"page": loc_slug, "section": title})
            out.append(render_section(key, title, "gap", f"Not documented for this location at any level."))
        else:
            if source == "gap":
                gaps_log.append({"page": loc_slug, "section": title})
            out.append(render_section(key, title, source, html))
    return out


def build_tab_tiles(loc_text, tab_tiles, gaps_log, loc_slug):
    """Dashboard: one tile per schema entry, all sourced from the location
    .md file. Used for every tab except Center Information on the dashboard
    (which mixes in the ops CSV — see build_center_info_tab)."""
    tiles, any_gap = [], False
    for key, title, heading in tab_tiles:
        html, source = resolve_tile_content(loc_text, heading)
        if html is None:
            gaps_log.append({"page": loc_slug, "section": title})
            tiles.append(dashboard_tile(key, title, "gap", "Not documented for this location at any level. Flag before an agent answers this for a customer."))
            any_gap = True
        else:
            if source == "gap":
                gaps_log.append({"page": loc_slug, "section": title})
                any_gap = True
            tiles.append(dashboard_tile(key, title, source, html))
    return tiles, any_gap


def build_location_page(loc, gaps_log):
    loc_text = read(os.path.join(BASE_DIR, "location-knowledge", loc["file"]))

    body = [f'<h1>{loc["name"]}</h1><p>Paw Sweet Paw · Irvine, CA · Center {loc["center_id"]}</p>'
            f'<p><a href="{loc["slug"]}-dashboard.html">Open the Dashboard View (at-a-glance, no scrolling) →</a></p>']

    all_tiles = [(key, title) for _, _, tiles in FULL_SCHEMA for key, title, _ in tiles]
    body.append(build_toc(all_tiles))

    for _, tab_title, tab_tiles in FULL_SCHEMA:
        body.append(f'<h2 class="group-header">{tab_title}</h2>')
        body.extend(build_tab_sections(loc_text, tab_tiles, gaps_log, loc["slug"]))

    identity_block = extract_h3_block(loc_text, "Contact Info") or ""
    address = extract_field(identity_block, "Address") or ""
    phone = extract_field(identity_block, "Direct center number") or ""
    jsonld = f"""<script type="application/ld+json">
{json.dumps({
    "@context": "https://schema.org",
    "@type": "LocalBusiness",
    "name": loc["name"],
    "url": f"{SITE_BASE_URL}/psp/{loc['slug']}.html",
    "address": address,
    "telephone": phone,
    "brand": "Paw Sweet Paw",
}, indent=2)}
</script>"""

    return PAGE_TEMPLATE.format(
        title=loc["name"],
        description=f'Rates, hours, policies, and booking info for {loc["name"]}, part of the Paw Sweet Paw brand.',
        canonical=f"{SITE_BASE_URL}/psp/{loc['slug']}.html",
        jsonld=jsonld,
        nav=build_nav("psp"),
        body="\n".join(body),
    )


# ── Center Dashboard (frontend, at-a-glance) ──────────────────────────────
# Same schema as the backend page, organized into tabs instead of a single
# scrolling document. Every component (tile) has a FIXED size — same
# width/height on every location's dashboard, same position within its tab,
# whether the tile is packed with tables or is a two-line gap notice.
DASHBOARD_TEMPLATE = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title} — Dashboard</title>
<meta name="robots" content="noindex, follow">
<style>
  :root {{
    color-scheme: light dark;
    --ink: #1c1c1e; --bg: #f6f7f9; --panel: #ffffff; --border: #e6e8ee;
    --muted: #767a85; --accent: #0a6cff; --accent-soft: #e8eefc; --accent-ink: #24478f;
    --radius: 14px; --radius-sm: 8px;
    --shadow: 0 1px 2px rgba(16,24,40,.04), 0 4px 14px rgba(16,24,40,.05);
    --scroll-shadow: rgba(0,0,0,.18);
    --tile-h: 236px; --cols: 3;
  }}
  @media (prefers-color-scheme: dark) {{
    :root {{ --ink: #e8e8ea; --bg: #121317; --panel: #1b1c22; --border: #2c2d36; --muted: #9498a3; --accent-soft: #16213f; --accent-ink: #8fb0f7; --shadow: 0 1px 2px rgba(0,0,0,.3), 0 4px 14px rgba(0,0,0,.25); --scroll-shadow: rgba(0,0,0,.55); }}
  }}
  /* Reduced desktop window (e.g. a laptop split-screened, not maximized): drop to 2 columns, keep everything else identical. */
  @media (max-width: 1180px) {{ :root {{ --cols: 2; --tile-h: 220px; }} }}
  /* Phone/narrow tablet: stack to 1 column. Desktop stays the reference layout; this tier only needs to stay usable. */
  @media (max-width: 720px) {{ :root {{ --cols: 1; --tile-h: 210px; }} }}

  * {{ box-sizing: border-box; }}
  html, body {{ height: 100%; margin: 0; }}
  body {{ display: flex; flex-direction: column; font-family: -apple-system, "Segoe UI", Roboto, sans-serif; color: var(--ink); background: var(--bg); overflow: hidden; }}
  @media (max-width: 720px) {{ body {{ overflow-y: auto; }} }}

  .dash-header {{ flex: 0 0 auto; display: flex; align-items: center; justify-content: space-between; gap: 16px; padding: 16px 28px; background: #1c2b4a; color: #fff; flex-wrap: wrap; }}
  .dash-header h1 {{ font-size: 1.15rem; margin: 0; letter-spacing: -0.01em; }}
  .dash-header .sub {{ font-size: 0.8rem; color: #aebbdc; margin-top: 2px; }}
  .dash-header a {{ color: #cfe0ff; font-size: 0.82rem; text-decoration: none; white-space: nowrap; }}
  .dash-header a:hover {{ text-decoration: underline; }}
  .header-links {{ display: flex; align-items: center; gap: 16px; flex-wrap: wrap; }}
  .pos-booking-btn {{ display: inline-flex; align-items: center; gap: 8px; background: #fff; color: #1c2b4a !important; padding: 6px 14px; border-radius: 999px; font-size: 0.82rem; font-weight: 600; text-decoration: none !important; box-shadow: 0 1px 3px rgba(0,0,0,.2); }}
  .pos-booking-btn:hover {{ background: #eef1f8; }}
  .pos-booking-btn img {{ height: 24px; width: auto; display: block; }}

  .tabs {{ flex: 0 0 auto; display: flex; gap: 4px; padding: 0 28px; background: var(--panel); border-bottom: 1px solid var(--border); overflow-x: auto; }}
  .tab-btn {{ appearance: none; border: none; background: transparent; color: var(--muted); font-size: 0.86rem; font-weight: 600; padding: 14px 6px; margin: 0 12px; cursor: pointer; border-bottom: 2.5px solid transparent; transition: color .15s, border-color .15s; white-space: nowrap; }}
  .tab-btn:first-child {{ margin-left: 0; }}
  .tab-btn:hover {{ color: var(--ink); }}
  .tab-btn.active {{ color: var(--accent); border-bottom-color: var(--accent); }}
  .tab-btn .flag {{ display: inline-block; margin-left: 6px; width: 7px; height: 7px; border-radius: 50%; background: #d9534f; vertical-align: middle; }}

  .content {{ flex: 1 1 auto; min-height: 0; overflow-y: auto; padding: 24px 28px 48px; }}
  .panel {{ display: none; }}
  .panel.active {{ display: block; }}

  /* Fixed-size tile grid: --cols columns, --tile-h row height. A tile's
     (cols,rows) span is set inline per-component and stays constant across
     every location's dashboard — only the number of grid columns adapts
     by breakpoint. Content that doesn't fill a tile just leaves it empty;
     content that overflows scrolls inside the tile, never resizes it.
     `dense` packing lets a later, narrower tile backfill a gap left by an
     earlier wide/tall one (e.g. a 2-wide tile after two 1-wide tiles would
     otherwise strand an empty cell and push the whole tab a row taller
     than it needs to be — dense closes that gap automatically). */
  .dash-grid {{ display: grid; grid-template-columns: repeat(var(--cols), 1fr); grid-auto-rows: var(--tile-h); grid-auto-flow: dense; gap: 16px; }}
  .tile {{
    grid-column: span min(var(--tcols, 1), var(--cols));
    grid-row: span var(--trows, 1);
    position: relative;
    border: 1px solid var(--border); box-shadow: var(--shadow);
    border-radius: var(--radius); padding: 16px 18px; overflow-y: auto;
    /* Scroll-shadow trick: a "cover" gradient (matches the panel color,
       scrolls WITH the content via background-attachment:local) sits
       exactly over a "shadow" gradient (fixed to the box via
       background-attachment:scroll). At rest — nothing to scroll, or
       already scrolled all the way — the cover fully hides the shadow.
       As soon as there's more content in that direction, the cover slides
       out of the way and the shadow becomes visible. No JS needed, and it
       tracks scroll position live. */
    background-color: var(--panel);
    background-image:
      linear-gradient(var(--panel) 30%, transparent),
      linear-gradient(transparent, var(--panel) 70%),
      radial-gradient(farthest-side at 50% 0, var(--scroll-shadow), transparent),
      radial-gradient(farthest-side at 50% 100%, var(--scroll-shadow), transparent);
    background-repeat: no-repeat;
    background-size: 100% 28px, 100% 28px, 100% 12px, 100% 12px;
    background-position: top, bottom, top, bottom;
    background-attachment: local, local, scroll, scroll;
  }}
  .scroll-hint {{ display: none; position: absolute; left: 0; right: 0; bottom: 6px; justify-content: center; pointer-events: none; }}
  .tile.has-overflow .scroll-hint {{ display: flex; }}
  .scroll-hint span {{ background: var(--accent-soft); color: var(--accent-ink); font-size: 0.66rem; font-weight: 700; letter-spacing: 0.02em; padding: 3px 10px; border-radius: 999px; box-shadow: 0 1px 4px rgba(0,0,0,.18); }}
  .tile h2 {{ font-size: 0.98rem; margin: 2px 0 8px; }}
  .tile h3 {{ font-size: 0.86rem; margin: 1em 0 4px; }}
  .tile h3:first-of-type {{ margin-top: 0.2em; }}
  .tile table {{ border-collapse: collapse; width: 100%; font-size: 0.83rem; margin: 0.5em 0; }}
  .tile th, .tile td {{ border: 1px solid var(--border); padding: 5px 9px; text-align: left; }}
  .tile th {{ background: var(--bg); position: sticky; top: 0; }}
  .tile ul {{ margin: 0.35em 0; padding-left: 1.2em; line-height: 1.5; }}
  .tile li {{ margin-bottom: 2px; }}
  .tile p {{ margin: 0.45em 0; line-height: 1.5; }}
  .source-banner {{ display: inline-block; font-size: 0.66rem; font-weight: 700; letter-spacing: 0.04em; text-transform: uppercase; padding: 3px 8px; border-radius: 5px; margin: 0 0 8px; }}
  .source-location {{ background: #e6f4ea; color: #1a7431; }}
  .source-brand {{ background: var(--accent-soft); color: var(--accent-ink); }}
  .source-default {{ background: #f4eee3; color: #7a5b1e; }}
  .source-gap {{ background: #fbe7e7; color: #9c2b2b; }}
  .source-system {{ background: #ece9f7; color: #4d3d99; }}
  @media (prefers-color-scheme: dark) {{
    .source-location {{ background: #10331b; color: #6fd98c; }}
    .source-default {{ background: #332a15; color: #e0bb6a; }}
    .source-gap {{ background: #3a1717; color: #f0a3a3; }}
    .source-system {{ background: #2a2440; color: #b8a8f0; }}
  }}
  .gap-note {{ font-style: italic; color: #9c2b2b; }}
  @media (prefers-color-scheme: dark) {{ .gap-note {{ color: #f0a3a3; }} }}
  .reference-note {{ color: var(--muted); font-size: 0.85rem; }}
  .conflict-note {{ background: #fff4e5; border: 1px solid #f0d9a8; border-radius: 8px; padding: 8px 12px; font-size: 0.8rem; margin: 8px 0; color: #7a5b1e; }}
  @media (prefers-color-scheme: dark) {{ .conflict-note {{ background: #332a15; border-color: #4a3d20; color: #e0bb6a; }} }}
  .kv {{ margin: 0.3em 0; font-size: 0.85rem; line-height: 1.5; }}
  .kv b {{ font-weight: 600; }}
  .toc-links {{ display: flex; flex-wrap: wrap; gap: 8px; margin-top: 10px; }}
  .toc-links a {{ display: inline-block; color: var(--accent-ink); background: var(--accent-soft); text-decoration: none; padding: 5px 12px; border-radius: 999px; font-size: 0.8rem; }}
  .toc-links a.gap-link {{ color: #9c2b2b; background: #fbe7e7; }}
  @media (prefers-color-scheme: dark) {{ .toc-links a.gap-link {{ color: #f0a3a3; background: #3a1717; }} }}
</style>
</head>
<body>
  <div class="dash-header">
    <div>
      <h1>{title}</h1>
      <div class="sub">Paw Sweet Paw · Center {center_id} · Dashboard view</div>
    </div>
    <div class="header-links">
      {pos_booking_button}
      <a href="{slug}.html">View Full Knowledgebase Page →</a>
    </div>
  </div>

  <div class="tabs">{tab_buttons}</div>

  <div class="content">{panels}</div>

<script>
(function() {{
  var buttons = document.querySelectorAll('.tab-btn');
  var panels = document.querySelectorAll('.panel');
  // A tile's real scrollHeight can only be measured once its panel is
  // actually laid out (display:none panels report 0), so this runs every
  // time a panel becomes active, plus again on resize since the 2-col/
  // 1-col breakpoints change --tile-h and can flip a tile's overflow state.
  function markOverflowingTiles(panel) {{
    panel.querySelectorAll('.tile').forEach(function(t) {{
      t.classList.toggle('has-overflow', t.scrollHeight > t.clientHeight + 1);
    }});
  }}
  function activate(key) {{
    buttons.forEach(function(b) {{ b.classList.toggle('active', b.dataset.tab === key); }});
    panels.forEach(function(p) {{
      var isActive = p.dataset.tab === key;
      p.classList.toggle('active', isActive);
      if (isActive) markOverflowingTiles(p);
    }});
  }}
  buttons.forEach(function(b) {{
    b.addEventListener('click', function() {{
      activate(b.dataset.tab);
      history.replaceState(null, '', '#' + b.dataset.tab);
    }});
  }});
  window.addEventListener('resize', function() {{
    var activePanel = document.querySelector('.panel.active');
    if (activePanel) markOverflowingTiles(activePanel);
  }});
  var initial = (location.hash || '').replace('#', '') || buttons[0].dataset.tab;
  if (![].some.call(buttons, function(b) {{ return b.dataset.tab === initial; }})) {{ initial = buttons[0].dataset.tab; }}
  activate(initial);
}})();
</script>
</body>
</html>
"""


def dashboard_tile(key, title, source_label, html_or_gap_text):
    cols, rows = TILE_SPANS.get(key, (1, 1))
    banner = f'<div class="source-banner {BANNER_CLASS[source_label]}">{BANNER_TEXT[source_label]}</div>'
    # A <div> wrapper (not <p>) — gap content is already-converted markdown
    # that may itself contain <p> tags, and nesting <p> inside <p> is invalid HTML.
    body = f'<div class="gap-note">{html_or_gap_text}</div>' if source_label == "gap" else html_or_gap_text
    # has-overflow is toggled by JS (see DASHBOARD_TEMPLATE's script) once the
    # tile's actual content height is known; the hint markup is always
    # present but only shown via that class, so a tile that fits needs no cue.
    scroll_hint = '<div class="scroll-hint"><span>Scroll for more ▾</span></div>'
    return f'<div class="tile" style="--tcols:{cols};--trows:{rows}">{banner}<h2>{title}</h2>{body}{scroll_hint}</div>'


def _kv(label, value):
    return f'<p class="kv"><b>{label}:</b> {value}</p>' if value else ""


def _gap_kv(label):
    return f'<p class="kv"><b>{label}:</b> <span class="gap-note">Not on file</span></p>'


def _link(url, text=None):
    url = (url or "").strip()
    if not url.startswith("http"):
        return url  # e.g. "No online booking currently" — not a real link
    return f'<a href="{url}" target="_blank" rel="noopener">{text or url}</a>'


def build_center_info_tab(loc, row, loc_text, gaps_log):
    """The one dashboard tab that mixes source types: Contact Info /
    Business Links / Hours come from the ops CSV (system, kept live outside
    the Markdown chain); Discounts, Organizational Chart, and Tours come
    from the location .md file. Organizational Chart is a placeholder for
    now — Center Manager and the rest of the org chart are intentionally
    deferred (see open-questions.md) rather than shown from a possibly-wrong
    source, so it's not treated as a gap or checked against the CSV."""
    if not row:
        gap = dashboard_tile(
            "contact-info", "Contact Info", "gap",
            f'No record found for Center {loc["center_id"]} in the center info CSV. Flag for follow-up.'
        )
        button = '<button class="tab-btn" data-tab="center-info">Center Information<span class="flag" title="Contains a gap"></span></button>'
        panel = f'<div class="panel" data-tab="center-info"><div class="dash-grid">{gap}</div></div>'
        return button, panel

    g = lambda field: (row.get(field) or "").strip()
    any_gap = False

    contact_html = "".join([
        _kv("Address", g("Full_Address")),
        _kv("Direct center number", g("Phone_Main")),
        _kv("MoeGo texting number", g("Phone_SMS_MoeGo")),
        _kv("Location email", g("Center_Generic_Email")),
        _gap_kv("Group Manager email"),
        _gap_kv("Call center number for location"),
    ])

    org_block = extract_h3_block(loc_text, "Organizational Chart") or ""
    org_html = convert(org_block) if org_block else '<p class="reference-note">Placeholder — see open-questions.md.</p>'

    discounts_block = extract_h3_block(loc_text, "Discounts Offered")
    if discounts_block:
        discounts_html = convert(discounts_block)
    else:
        gaps_log.append({"page": loc["slug"], "section": "Discounts Offered"})
        any_gap = True
        discounts_html = '<p class="gap-note">Not documented for this location.</p>'

    days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    hours_rows = "".join(f'<tr><td>{d}</td><td>{g(d) or "—"}</td></tr>' for d in days)
    hours_extra = extract_h3_block(loc_text, "Hours of Operation")
    hours_html = (
        '<table><thead><tr><th>Day</th><th>Hours</th></tr></thead><tbody>' + hours_rows + '</tbody></table>'
        + "".join([_kv("Special Hours", g("Special_Hours")), _kv("Pet Cameras", g("Pet Cameras"))])
        + (f'<h3>Check-in / Check-out Details</h3>{convert(hours_extra)}' if hours_extra else "")
    )

    pos_login_url = g("POS_Direct_Login_URL")
    pos_login_label = g("POS_Direct_Login") or "Log in"
    business_links = [
        ("Website", _link(g("Website_URL"))),
        ("Book Now (customers)", _link(g("Book Now Link"), "Book Now")),
        ("POS Direct Login (internal)", _link(pos_login_url, pos_login_label) if pos_login_url else g("POS_Direct_Login")),
        ("Google Business Profile", _link(g("GMB_Link"), "View listing")),
        ("Leave a Review", _link(g("GMB_LeaveAReview_Link"), "Review link")),
        ("Facebook", _link(g("Facebook_Link"), "Open")),
        ("Instagram", _link(g("Instagram_Link"), "Open")),
        ("TikTok", _link(g("TikTok_Link"), "Open")),
        ("YouTube", _link(g("YouTube_Link"), "Open")),
    ]
    links_html = "".join(_kv(label, val) for label, val in business_links if val) or '<p class="reference-note">No links on file.</p>'

    tours_block = extract_h3_block(loc_text, "Tour Information")
    if tours_block is None:
        gaps_log.append({"page": loc["slug"], "section": "Tour Information"})
        any_gap = True
        tours_html = '<p class="gap-note">Not documented for this location.</p>'
    else:
        gap = is_gap_text(tours_block)
        any_gap = any_gap or gap
        if gap:
            gaps_log.append({"page": loc["slug"], "section": "Tour Information"})
        tours_html = convert(tours_block)

    tiles = [
        dashboard_tile("contact-info", "Contact Info", "system", contact_html),
        dashboard_tile("org-chart", "Organizational Chart", "location", org_html),
        dashboard_tile("discounts", "Discounts Offered", "gap" if not discounts_block else "location", discounts_html),
        dashboard_tile("hours", "Hours of Operation", "system", hours_html),
        dashboard_tile("business-links", "Business Links", "system", links_html),
        dashboard_tile("tours", "Tour Information", "gap" if (tours_block is None or is_gap_text(tours_block)) else "location", tours_html),
    ]
    flag = '<span class="flag" title="Contains a gap"></span>' if any_gap else ""
    button = f'<button class="tab-btn" data-tab="center-info">Center Information{flag}</button>'
    panel = f'<div class="panel" data-tab="center-info"><div class="dash-grid">{"".join(tiles)}</div></div>'
    return button, panel


def build_dashboard_page(loc, center_info, gaps_log):
    loc_text = read(os.path.join(BASE_DIR, "location-knowledge", loc["file"]))

    tab_buttons, panels = [], []

    # Center Information is the leftmost tab — a CS rep who knows the layout
    # goes here first for anything ops/system related, then into the
    # service-specific tabs to its right.
    center_info_button, center_info_panel = build_center_info_tab(loc, center_info.get(loc["center_id"]), loc_text, gaps_log)
    tab_buttons.append(center_info_button)
    panels.append(center_info_panel)

    for tab_key, tab_title, tab_tiles in FULL_SCHEMA[1:]:
        tiles, any_gap = build_tab_tiles(loc_text, tab_tiles, gaps_log, loc["slug"])
        flag = '<span class="flag" title="Contains a gap"></span>' if any_gap else ""
        tab_buttons.append(f'<button class="tab-btn" data-tab="{tab_key}">{tab_title}{flag}</button>')
        panels.append(f'<div class="panel" data-tab="{tab_key}"><div class="dash-grid">{"".join(tiles)}</div></div>')

    return DASHBOARD_TEMPLATE.format(
        title=loc["name"],
        center_id=loc["center_id"],
        slug=loc["slug"],
        pos_booking_button=build_pos_booking_button(center_info.get(loc["center_id"])),
        tab_buttons="".join(tab_buttons),
        panels="".join(panels),
    )


OVERVIEW_SCHEMA = [
    ("locations", "Locations"), ("booking", "Website & Booking"), ("rates", "Shared Rates"),
    ("daycare-packages", "Shared Daycare Packages"), ("training", "Shared Training Pricing"),
    ("hours", "Shared Hours"), ("vaccines", "Shared Vaccine Requirements (Dogs)"),
    ("boarding-together", "Shared Boarding Together Rules"), ("grooming", "Shared Grooming Service Names & Pricing"),
    ("addons", "Shared Add-Ons"), ("differences", "Where the Two Locations Differ"),
]


def build_overview_page(brand_text, gaps_log):
    body = ['<h1>Paw Sweet Paw</h1><p>Two locations in Irvine, CA.</p>']
    body.append(build_toc(OVERVIEW_SCHEMA))
    for anchor_key, heading in OVERVIEW_SCHEMA:
        block = extract_h2_block(brand_text, heading)
        if block:
            body.append(render_section(anchor_key, heading, "brand", convert(block)))
        else:
            gaps_log.append({"page": "index", "section": heading})
            body.append(render_section(anchor_key, heading, "gap", "Not on file at the brand level."))

    body.append(build_knowledgebase_map())

    return PAGE_TEMPLATE.format(
        title="Overview",
        description="Brand-wide rates, hours, and policies shared by both Paw Sweet Paw locations, plus a full map of the knowledgebase.",
        canonical=f"{SITE_BASE_URL}/psp/index.html",
        jsonld="",
        nav=build_nav("psp"),
        body="\n".join(body),
    )


# Company-wide policy topics shown on the shared best-friends-standard.html
# reference page (linked to from each location's "BFPC Policy Docs" tile).
COMPANY_TOPICS = [
    ("what-owners-can-bring", "What Owners Can Bring", "What Owners Can Bring"),
    ("food", "Food", "Food"),
    ("drop-off", "Drop-Off", "Drop-Off"),
    ("medication-policy", "Medication Policy", "Medication Policy"),
    ("shared-accommodations", "Shared Accommodations", "Shared Accommodations"),
    ("animals-we-cannot-accept", "Animals We Cannot Accept", "Animals We Cannot Accept"),
    ("parasites", "Parasites", "Parasites"),
    ("complaints-concerns", "Complaints and Concerns", "Complaints & Concerns"),
]


def build_standards_page(defaults_text, gaps_log):
    body = ['<h1>Best Friends Standard</h1>'
            '<p>Company-wide policy that applies across every Best Friends Pet Care brand and location, unless a '
            'brand or location file says otherwise. This is the single copy — brand and location pages link here '
            'by topic instead of repeating the text.</p>']
    body.append(build_toc([(key, display) for key, heading, display in COMPANY_TOPICS]))
    for anchor_key, heading, display in COMPANY_TOPICS:
        block = extract_h2_block(defaults_text, heading) or extract_h3_block(defaults_text, heading)
        if block:
            body.append(render_section(anchor_key, display, "default", convert(block)))
        else:
            gaps_log.append({"page": "best-friends-standard", "section": display})
            body.append(render_section(anchor_key, display, "gap", "Not documented at the company level."))

    return PAGE_TEMPLATE.format(
        title="Best Friends Standard",
        description="Company-wide policy shared across all Best Friends Pet Care brands and locations.",
        canonical=f"{SITE_BASE_URL}/best-friends-standard.html",
        jsonld="",
        nav=build_nav("root"),
        body="\n".join(body),
    )


def build_knowledgebase_map():
    """Human-readable index of every page and every section anchor — for
    someone who knows the system to jump straight to an exact fact, and to
    make clear what topics exist (or don't) across the pilot."""
    rows = ['<section id="knowledgebase-map"><div class="source-banner source-brand">Navigation</div>'
            '<h2>Full Knowledgebase Map</h2>'
            '<p>Every page and section in the Paw Sweet Paw pilot, plus the shared company standard. Link directly to any row.</p>']

    standards_items = [(key, display) for key, heading, display in COMPANY_TOPICS]
    rows.append(f'<p><strong>Best Friends Standard</strong> (<a href="../best-friends-standard.html">best-friends-standard.html</a>)<br>'
                + " · ".join(f'<a href="../best-friends-standard.html#{key}">{title}</a>' for key, title in standards_items)
                + '</p>')

    all_tiles = [(key, title) for _, _, tiles in FULL_SCHEMA for key, title, _ in tiles]
    pages = [("index.html", "Overview", OVERVIEW_SCHEMA)] + [
        (f'{loc["slug"]}.html', loc["name"], all_tiles) for loc in LOCATIONS
    ]
    for filename, page_title, items in pages:
        links = " · ".join(f'<a href="{filename}#{key}">{title}</a>' for key, title in items)
        rows.append(f'<p><strong>{page_title}</strong> (<a href="{filename}">{filename}</a>)<br>{links}</p>')
    rows.append("</section>")
    return "\n".join(rows)


def write_sitemap(psp_files):
    urls = [f"  <url><loc>{SITE_BASE_URL}/best-friends-standard.html</loc></url>"]
    urls += [f"  <url><loc>{SITE_BASE_URL}/psp/{f}</loc></url>" for f in psp_files]
    xml = f'<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n' + "\n".join(urls) + "\n</urlset>\n"
    with open(os.path.join(SITE_ROOT, "sitemap.xml"), "w", encoding="utf-8") as f:
        f.write(xml)


def write_robots():
    content = f"User-agent: *\nAllow: /\nSitemap: {SITE_BASE_URL}/sitemap.xml\n"
    with open(os.path.join(SITE_ROOT, "robots.txt"), "w", encoding="utf-8") as f:
        f.write(content)


def main():
    os.makedirs(PSP_DIR, exist_ok=True)
    gaps_log = []

    defaults_text = read(os.path.join(BASE_DIR, "company-defaults.md"))
    brand_text = read(os.path.join(BASE_DIR, "brand-knowledge", "paw-sweet-paw-brand.md"))
    center_info = load_center_info()
    copy_pos_logo_assets()

    root_pages = {"best-friends-standard.html": build_standards_page(defaults_text, gaps_log)}

    psp_pages = {"index.html": build_overview_page(brand_text, gaps_log)}
    for loc in LOCATIONS:
        psp_pages[f'{loc["slug"]}.html'] = build_location_page(loc, gaps_log)
        psp_pages[f'{loc["slug"]}-dashboard.html'] = build_dashboard_page(loc, center_info, gaps_log)

    for filename, html in root_pages.items():
        path = os.path.join(SITE_ROOT, filename)
        with open(path, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"  wrote {path}")

    for filename, html in psp_pages.items():
        path = os.path.join(PSP_DIR, filename)
        with open(path, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"  wrote {path}")

    write_sitemap([f for f in psp_pages.keys() if not f.endswith("-dashboard.html")])
    write_robots()
    print(f"  wrote {os.path.join(SITE_ROOT, 'sitemap.xml')}")
    print(f"  wrote {os.path.join(SITE_ROOT, 'robots.txt')}")

    # The backend knowledgebase page and the dashboard page walk the same
    # schema, so the same (page, section) gap gets logged from both builders —
    # dedupe before writing so gaps.json (and any future missing-info-bank
    # frequency count built on top of it) isn't artificially doubled.
    seen = set()
    deduped_gaps = []
    for gap in gaps_log:
        key = (gap["page"], gap["section"])
        if key not in seen:
            seen.add(key)
            deduped_gaps.append(gap)

    gaps_path = os.path.join(SITE_ROOT, "gaps.json")
    with open(gaps_path, "w", encoding="utf-8") as f:
        json.dump(deduped_gaps, f, indent=2)
    print(f"  wrote {gaps_path} ({len(deduped_gaps)} gap(s) found)")

    print(f"\nDone. {len(root_pages) + len(psp_pages)} pages generated in {SITE_ROOT}")


if __name__ == "__main__":
    main()
