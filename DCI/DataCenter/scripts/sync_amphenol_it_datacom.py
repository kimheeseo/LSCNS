#!/usr/bin/env python3
from __future__ import annotations

import hashlib
import json
import math
import re
import shutil
import time
import unicodedata
from collections import Counter, defaultdict
from pathlib import Path
from urllib.parse import urljoin, urlparse

import requests

BASE = "https://www.amphenol.com/markets/it-datacom"
OFFICIAL_BASE = BASE
READER = "https://r.jina.ai/"
PAGE_SIZE = 50
ROOT = Path("DCI/DataCenter/product_catalog")
AMPHENOL_ROOT = ROOT / "Amphenol"
INDEX_PATH = ROOT / "bom-catalog-index.json"
CHECKED = "2026-10-05"

session = requests.Session()
session.headers.update({
    "User-Agent": "LSCNS-DataCenter-Catalog/1.0",
    "Accept": "text/plain,*/*;q=0.8",
    "X-No-Cache": "true",
    "X-Cache-Tolerance": "0",
})

def clean_text(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()

def fetch(url: str) -> str:
    reader_url = READER + url
    last = None
    for attempt in range(5):
        try:
            r = session.get(reader_url, timeout=120)
            r.raise_for_status()
            text = r.text
            if "IT Datacom" not in text:
                raise RuntimeError("Unexpected Reader response")
            return text
        except Exception as exc:
            last = exc
            time.sleep(3 + attempt * 3)
    raise RuntimeError(f"Failed to fetch {url} through Jina Reader: {last}")

def advertised_total(text: str) -> int:
    m = re.search(r"Showing\s+\d+\s+to\s+\d+\s+of\s+([\d,]+)\s+items", text, re.I)
    if not m:
        raise RuntimeError("Could not find Amphenol advertised item count")
    return int(m.group(1).replace(",", ""))

def _clean_md_title(line: str) -> str:
    line = clean_text(line)
    line = re.sub(r"^#+\s*", "", line)
    line = re.sub(r"^[-*]\s+", "", line)
    line = re.sub(r"^\*\*(.*?)\*\*$", r"\1", line)
    # Preserve markdown link labels (especially the Amphenol business name) while dropping URLs.
    line = re.sub(r"!\[[^\]]*\]\([^)]*\)", "", line)
    line = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", line)
    return clean_text(line)

def parse_page(text: str, page: int) -> list[dict]:
    page_url = f"{OFFICIAL_BASE}?PageSize={PAGE_SIZE}&pagenumber={page}"
    lines = [clean_text(x) for x in text.splitlines()]
    rows = []
    seen = Counter()

    # Reader preserves product image alt text as "Image: Product <name>" or markdown image alt.
    candidates = []
    for i, line in enumerate(lines):
        if not line:
            continue
        m = re.search(r"(?:Image:\s*Product\s+|!\[[^\]]*?Product\s+)(.+?)(?:\]\([^)]*\)|\]$|$)", line, re.I)
        if m:
            title = _clean_md_title(m.group(1))
            if title:
                candidates.append((i, title))

    # Some official listing rows have no image. Recover them from blocks between separators/business links.
    # Candidate product titles are non-navigation lines immediately followed by description text and before an
    # Amphenol business line. We only add these if the image-based count is short for the page.
    expected = PAGE_SIZE
    if page * PAGE_SIZE > advertised_total(text):
        expected = advertised_total(text) - (page - 1) * PAGE_SIZE

    for i, title in candidates:
        key = title.casefold()
        seen[key] += 1
        # Description: first substantive non-navigation line after the repeated title/image line.
        desc = ""
        business = "Amphenol"
        for j in range(i + 1, min(len(lines), i + 14)):
            x = _clean_md_title(lines[j])
            if not x or x == title or x.startswith("http") or "Image: Product" in x:
                continue
            if re.search(r"^(Amphenol|SV Microwave|Positronic|LUTZE|TPC Wire|EBY|Piher|Assembletech|RFS Technologies)", x, re.I):
                business = x[:160]
                continue
            if "Showing " in x or x in {"X", "Products", "Markets", "Businesses", "Sustainability", "Investors"}:
                continue
            if len(x) >= 18 and not desc:
                desc = x[:1200]
        rows.append({
            "name": title,
            "description": desc or "Official Amphenol IT Datacom listing product.",
            "business": business,
            "businessUrl": OFFICIAL_BASE,
            "sourcePage": page,
            "sourceUrl": page_url,
        })

    if len(rows) < expected:
        # Block fallback. Reader separates most products with horizontal rules. Parse likely title + description pairs.
        blocks = re.split(r"\n\s*(?:\* \* \*|---+)\s*\n", text)
        existing = Counter(r["name"].casefold() for r in rows)
        for block in blocks:
            blines = [_clean_md_title(x) for x in block.splitlines() if clean_text(x)]
            if len(blines) < 2:
                continue
            # Remove obvious page/navigation prose and markdown images.
            blines = [x for x in blines if x and "Image: Product" not in x and not x.startswith("![") and "Showing " not in x]
            if len(blines) < 2:
                continue
            # Product name is generally the first short standalone line in a product block.
            title = None
            for x in blines[:6]:
                if x in {"IT Datacom","Products","Markets","Businesses","Sustainability","Investors"}:
                    continue
                if len(x) <= 180 and not x.startswith("With our industry") and not x.startswith("Our primary"):
                    title = x
                    break
            if not title or existing[title.casefold()] > 0:
                continue
            # Require a later Amphenol-business marker to avoid navigation/footer blocks.
            business = next((x for x in blines if re.search(r"^(Amphenol|SV Microwave|Positronic|LUTZE|TPC Wire|EBY|Piher|Assembletech|RFS Technologies)", x, re.I)), None)
            if not business:
                continue
            desc = next((x for x in blines[1:] if x != title and x != business and len(x) >= 18), "")
            rows.append({
                "name": title,
                "description": desc[:1200] or "Official Amphenol IT Datacom listing product.",
                "business": business[:160],
                "businessUrl": OFFICIAL_BASE,
                "sourcePage": page,
                "sourceUrl": page_url,
            })
            existing[title.casefold()] += 1
            if len(rows) >= expected:
                break

    if len(rows) != expected:
        preview = [r["name"] for r in rows[:5]]
        raise RuntimeError(f"Page {page}: parsed {len(rows)} products, expected {expected}; preview={preview}")
    return rows

def classify(name: str, description: str) -> str:
    s = (name + " " + description).lower()
    checks = [
        ("Optical Transceivers and AOC", r"transceiver|active optical|\baoc\b|optical module"),
        ("Power Distribution Panels", r"fuse panel|circuit breaker panel|power distribution panel|breaker panel"),
        ("High-Speed Cable Assemblies", r"cable assembl|\bdac\b|direct attach|overpass|omni-path|copper cable"),
        ("Power Connectors and Busbar", r"barklip|busbar|power connector|energyedge|radsok|battery connector|power cable|power edge"),
        ("Storage and PCIe-SAS Interconnects", r"pcie|pci express|\bsas\b|oculink|slimsas|minisas|sata|edsff|u\.2|u\.3|sff-"),
        ("Memory and Card Edge Connectors", r"memory module|\bdimm\b|so-dimm|camm|ddr\d|lpddr|card edge"),
        ("Backplane and Orthogonal Connectors", r"backplane|orthogonal|crossbow|metral|airmax|hard metric"),
        ("High-Speed Board Connectors", r"board-to-board|mezzanine|bergstak|cstack|interposer|btb"),
        ("Wire-to-Board and FFC-FPC", r"ffc|fpc|wire-to-board|wire to board|agillink|dubox"),
        ("Fiber Optic Connectivity", r"fiber optic|fibre optic|\bmtp\b|\bmpo\b|\blc\b connector|optik"),
        ("Ethernet USB and External I-O", r"rj45|ethernet|\busb\b|displayport|hdmi|external i/o|type-c"),
        ("RF and Coaxial Connectivity", r"\brf\b|coax|sma|smpm?|n-type|bnc|tnc|vna|microwave"),
        ("Rugged Circular and D-Sub", r"circular|\bm12\b|\bm8\b|d-sub|micro-d|arinc|mil-|ip67|ip68|ip69"),
        ("Terminal Blocks and General Interconnect", r"terminal block|barrier strip|header|receptacle|socket|connector"),
    ]
    for label, pattern in checks:
        if re.search(pattern, s, re.I):
            return label
    return "Sensors Materials and Other"

def extract_key_spec(name: str, desc: str) -> str:
    text = name + " " + desc
    specs = []
    patterns = [
        r"\b\d+(?:\.\d+)?\s*(?:Tb/s|Gb/s|Gbps|GT/s|GHz)\b",
        r"\bPCIe(?:®)?\s*Gen\s*\d(?:\.\d)?\b",
        r"\b(?:up to\s*)?\d+(?:\.\d+)?\s*A\b",
        r"\b\d+(?:\.\d+)?\s*mm\s*pitch\b",
        r"\b\d+\s*(?:F|fiber|fibers)\b",
        r"\b(?:OSFP|QSFP(?:28|-DD)?|SFP\+?|MPO|MTP|LC|RJ45|USB(?:\s*Type-?C)?)\b",
    ]
    for pattern in patterns:
        for match in re.findall(pattern, text, re.I):
            value = clean_text(match if isinstance(match, str) else " ".join(match))
            if value and value.lower() not in {x.lower() for x in specs}:
                specs.append(value)
            if len(specs) >= 4:
                return " · ".join(specs)
    return "See official IT Datacom description"

def infer_application(desc: str) -> str:
    s = desc.lower()
    labels = []
    for label, pattern in [
        ("Data center", r"data center|datacenter"),
        ("Server", r"server"),
        ("Storage", r"storage"),
        ("Networking", r"network"),
        ("HPC", r"high-performance computing|\bhpc\b"),
        ("Telecom", r"telecom|wireless"),
        ("Industrial", r"industrial"),
        ("AI", r"artificial intelligence|\bai\b|gpu"),
    ]:
        if re.search(pattern, s):
            labels.append(label)
    return " / ".join(labels[:4]) if labels else "IT Datacom"

def slugify(name: str) -> str:
    s = unicodedata.normalize("NFKD", name)
    s = s.encode("ascii", "ignore").decode("ascii").lower()
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    if not s:
        s = "product-" + hashlib.sha1(name.encode("utf-8")).hexdigest()[:10]
    return s[:96]

def write_catalogs(items: list[dict], total: int):
    if AMPHENOL_ROOT.exists():
        shutil.rmtree(AMPHENOL_ROOT)
    AMPHENOL_ROOT.mkdir(parents=True, exist_ok=True)

    grouped = defaultdict(list)
    for row in items:
        row["category"] = classify(row["name"], row["description"])
        row["keySpec"] = extract_key_spec(row["name"], row["description"])
        row["application"] = infer_application(row["description"])
        grouped[row["category"]].append(row)

    source_rows = []
    for category in sorted(grouped):
        rows = grouped[category]
        folder = AMPHENOL_ROOT / category
        folder.mkdir(parents=True, exist_ok=True)
        products = {}
        used = Counter()
        source_urls = []
        for idx, row in enumerate(rows, 1):
            key = slugify(row["name"])
            used[key] += 1
            if used[key] > 1:
                key = f"{key}--{used[key]}"
            source_urls.append(row["sourceUrl"])
            products[key] = {
                "name": row["name"],
                "description": row["description"],
                "officialUrl": row["sourceUrl"],
                "businessUrl": row["businessUrl"],
                "specs": {
                    "Product Type": category,
                    "Key Spec": row["keySpec"],
                    "Application": row["application"],
                    "Amphenol Business": row["business"],
                    "Source Page": f"IT Datacom p.{row['sourcePage']}",
                },
            }
            source_rows.append({
                "name": row["name"],
                "category": category,
                "business": row["business"],
                "description": row["description"],
                "sourcePage": row["sourcePage"],
                "sourceUrl": row["sourceUrl"],
            })
        manifest = {
            "schemaVersion": 1,
            "company": "Amphenol",
            "category": category,
            "displayName": category,
            "description": f"Amphenol IT Datacom products classified for the DataCenter Tool: {category}. Classification is tool-side; product names/descriptions come from the official IT Datacom listing.",
            "checked": CHECKED,
            "officialUrl": OFFICIAL_BASE,
            "sourceUrls": sorted(set(source_urls)),
            "sourceCount": len(rows),
            "sourceTotal": total,
            "storageMode": "metadata only; official listing URLs retained",
            "comparisonFields": ["Product Type", "Key Spec", "Application", "Amphenol Business", "Source Page"],
            "products": products,
        }
        (folder / "catalog.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        (folder / "readme.md").write_text(
            f"# Amphenol {category}\n\n"
            f"- Official market: {BASE}\n"
            f"- Products in this tool category: {len(rows)}\n"
            f"- Official IT Datacom total at sync: {total}\n"
            f"- Reviewed: {CHECKED}\n\n"
            "Product names and descriptions are sourced from Amphenol's official IT Datacom market listing. "
            "The category grouping is a DataCenter Tool classification for comparison/BOM use.\n",
            encoding="utf-8",
        )

    root_summary = {
        "schemaVersion": 1,
        "company": "Amphenol",
        "officialUrl": OFFICIAL_BASE,
        "checked": CHECKED,
        "sourceTotal": total,
        "parsedTotal": len(items),
        "categoryCounts": dict(sorted(Counter(x["category"] for x in items).items())),
        "products": source_rows,
    }
    (AMPHENOL_ROOT / "it-datacom-source.json").write_text(
        json.dumps(root_summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    (AMPHENOL_ROOT / "readme.md").write_text(
        "# Amphenol IT Datacom\n\n"
        f"Official source: {BASE}\n\n"
        f"All {total} products visible in the official IT Datacom listing were synchronized on {CHECKED}. "
        "Products are split into DataCenter Tool comparison categories; source product names/descriptions are retained.\n",
        encoding="utf-8",
    )
    return grouped

def update_bom_index(grouped):
    if not INDEX_PATH.exists():
        return
    data = json.loads(INDEX_PATH.read_text(encoding="utf-8"))
    roles = data.setdefault("roles", {})
    # Remove prior Amphenol paths, then add the newly generated categories by BOM role.
    for role in list(roles):
        roles[role] = [p for p in roles[role] if not p.startswith("Amphenol/")]

    role_map = {
        "Optical Transceivers and AOC": ["optic"],
        "High-Speed Cable Assemblies": ["trunk", "patchCord"],
        "Fiber Optic Connectivity": ["connector", "adapter", "patchCord"],
        "High-Speed Board Connectors": ["connector"],
        "High-Speed I-O Connectors": ["connector"],
        "Storage and PCIe-SAS Interconnects": ["connector"],
        "Memory and Card Edge Connectors": ["connector"],
        "Backplane and Orthogonal Connectors": ["connector"],
        "Wire-to-Board and FFC-FPC": ["connector"],
        "Ethernet USB and External I-O": ["connector"],
        "RF and Coaxial Connectivity": ["connector"],
        "Rugged Circular and D-Sub": ["connector"],
        "Terminal Blocks and General Interconnect": ["connector"],
        "Power Connectors and Busbar": ["power"],
        "Power Distribution Panels": ["power"],
    }
    for category in grouped:
        rel = f"Amphenol/{category}/catalog.json"
        for role in role_map.get(category, []):
            roles.setdefault(role, []).append(rel)
    for role in roles:
        roles[role] = sorted(set(roles[role]))
    data["checked"] = CHECKED
    data["source"] = "DCI/DataCenter/product_catalog (업체별 부품 리스트)"
    data["totalStructuredCatalogs"] = len(list(ROOT.rglob("catalog.json")))
    data["amphenolItDatacom"] = {
        "officialUrl": OFFICIAL_BASE,
        "sourceTotal": sum(len(v) for v in grouped.values()),
        "categoryCount": len(grouped),
        "mode": "All official IT Datacom listing products indexed; tool-side category classification",
    }
    INDEX_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

def main():
    first_url = f"{BASE}?PageSize={PAGE_SIZE}&pagenumber=1"
    first_html = fetch(first_url)
    total = advertised_total(first_html)
    pages = math.ceil(total / PAGE_SIZE)

    items = []
    for page in range(1, pages + 1):
        html = first_html if page == 1 else fetch(f"{BASE}?PageSize={PAGE_SIZE}&pagenumber={page}")
        rows = parse_page(html, page)
        if not rows:
            raise RuntimeError(f"No products parsed from page {page}")
        items.extend(rows)
        time.sleep(0.25)

    # Preserve source listing cardinality. Exact name duplicates receive stable suffixed keys later.
    if len(items) != total:
        # Fallback to canonical 20-item pagination in case the site ignores PageSize=50.
        items = []
        page_size = 20
        pages = math.ceil(total / page_size)
        for page in range(1, pages + 1):
            html = fetch(f"{BASE}?PageSize={page_size}&pagenumber={page}")
            rows = parse_page(html, page)
            if not rows:
                raise RuntimeError(f"No products parsed from fallback page {page}")
            items.extend(rows)
            time.sleep(0.25)
    if len(items) != total:
        raise RuntimeError(f"Parsed {len(items)} products but Amphenol advertises {total}; refusing partial catalog")

    grouped = write_catalogs(items, total)
    update_bom_index(grouped)
    print(json.dumps({
        "officialTotal": total,
        "parsed": len(items),
        "categories": {k: len(v) for k, v in sorted(grouped.items())},
    }, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    main()
