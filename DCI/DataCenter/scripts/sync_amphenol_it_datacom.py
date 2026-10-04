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

from curl_cffi import requests
from bs4 import BeautifulSoup

BASE = "https://www.amphenol.com/markets/it-datacom"
PAGE_SIZE = 50
ROOT = Path("DCI/DataCenter/product_catalog")
AMPHENOL_ROOT = ROOT / "Amphenol"
INDEX_PATH = ROOT / "bom-catalog-index.json"
CHECKED = "2026-10-05"

session = requests.Session(impersonate="chrome")
session.headers.update({
    "User-Agent": "Mozilla/5.0 (compatible; LSCNS-DataCenter-Catalog/1.0; +https://github.com/kimheeseo/LSCNS)",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
})

def clean_text(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()

def fetch(url: str) -> str:
    last = None
    for attempt in range(5):
        try:
            r = session.get(url, timeout=60, allow_redirects=True)
            r.raise_for_status()
            if "IT Datacom" not in r.text:
                raise RuntimeError("Unexpected Amphenol response")
            return r.text
        except Exception as exc:
            last = exc
            time.sleep(2 + attempt * 2)
    raise RuntimeError(f"Failed to fetch {url}: {last}")

def advertised_total(html: str) -> int:
    text = BeautifulSoup(html, "html.parser").get_text(" ", strip=True)
    m = re.search(r"Showing\s+\d+\s+to\s+\d+\s+of\s+([\d,]+)\s+items", text, re.I)
    if not m:
        raise RuntimeError("Could not find Amphenol advertised item count")
    return int(m.group(1).replace(",", ""))

def card_for_image(img, title: str):
    node = img
    fallback = None
    for _ in range(10):
        node = getattr(node, "parent", None)
        if node is None:
            break
        try:
            imgs = node.find_all("img", alt=re.compile(r"^Product\s+", re.I))
            text = clean_text(node.get_text(" ", strip=True))
        except Exception:
            continue
        if len(imgs) == 1 and title.lower() in text.lower():
            fallback = node
            if len(text) >= len(title) + 35:
                return node
    return fallback or img.parent

def infer_business(card) -> tuple[str, str]:
    candidates = []
    for a in card.find_all("a", href=True):
        label = clean_text(a.get_text(" ", strip=True))
        href = urljoin(BASE, a.get("href", ""))
        host = urlparse(href).netloc.lower()
        if not label or "blob.core.windows.net" in host:
            continue
        if label in {"X", "1", "2", "3", "4", "5", "6", "7", "8", "9"}:
            continue
        if "amphenol" in label.lower() or "positronic" in label.lower() or "sv microwave" in label.lower() or "lutze" in label.lower() or "eby" in label.lower() or "piher" in label.lower():
            candidates.append((label, href))
    return candidates[-1] if candidates else ("Amphenol", BASE)

def extract_description(card, title: str, business: str) -> str:
    # Prefer paragraph-like blocks; otherwise use the compact card text.
    blocks = []
    for tag in card.find_all(["p", "div"]):
        text = clean_text(tag.get_text(" ", strip=True))
        if len(text) < 25 or text == title or text == business:
            continue
        if title in text and len(text) > 1200:
            continue
        blocks.append(text)
    if blocks:
        blocks.sort(key=lambda x: (title.lower() in x.lower(), len(x)))
        text = blocks[-1]
    else:
        text = clean_text(card.get_text(" ", strip=True))
    text = re.sub(re.escape(title), "", text, count=1, flags=re.I).strip(" -–—|:")
    if business:
        text = re.sub(r"\s*" + re.escape(business) + r"\s*$", "", text, flags=re.I)
    text = clean_text(text)
    return text[:1200]

def parse_page(html: str, page: int) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    page_url = f"{BASE}?PageSize={PAGE_SIZE}&pagenumber={page}"
    rows = []
    for img in soup.find_all("img"):
        alt = clean_text(img.get("alt", ""))
        m = re.match(r"^Product\s+(.+)$", alt, re.I)
        if not m:
            continue
        title = clean_text(m.group(1))
        if not title:
            continue
        card = card_for_image(img, title)
        business, business_url = infer_business(card)
        description = extract_description(card, title, business)
        rows.append({
            "name": title,
            "description": description,
            "business": business,
            "businessUrl": business_url,
            "sourcePage": page,
            "sourceUrl": page_url,
        })
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
            "officialUrl": BASE,
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
        "officialUrl": BASE,
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
        "officialUrl": BASE,
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
