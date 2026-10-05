#!/usr/bin/env python3
import json, re, time, html, hashlib
from pathlib import Path
from urllib.parse import urljoin
import requests
from bs4 import BeautifulSoup

BASE="https://www.amphenol.com/markets/it-datacom"
ROOT=Path("DCI/DataCenter/product_catalog/Amphenol")
EXPECTED=665
CHECKED="2026-10-05"
session=requests.Session()
session.headers.update({"User-Agent":"Mozilla/5.0 (compatible; LSCNS-DataCenter-Catalog/1.0)"})

CATEGORIES=[
 "Optical Transceivers and AOC",
 "Fiber Optic Connectivity",
 "High-Speed Cable Assemblies",
 "High-Speed Board Connectors",
 "Backplane and Orthogonal Connectors",
 "Memory and Card Edge Connectors",
 "Storage and PCIe-SAS Interconnects",
 "Ethernet USB and External I-O",
 "Wire-to-Board and FFC-FPC",
 "Power Connectors and Busbar",
 "Power Distribution Panels",
 "RF and Coaxial Connectivity",
 "Rugged Circular and D-Sub",
 "Terminal Blocks and General Interconnect",
 "Sensors Materials and Other",
]

def compact(s):
    return re.sub(r"\s+"," ",html.unescape(str(s or ""))).strip()

def slug(s):
    x=compact(s).lower()
    x=re.sub(r"[®™©]","",x)
    x=re.sub(r"[^a-z0-9]+","-",x).strip("-")
    return x[:110] or hashlib.sha1(s.encode()).hexdigest()[:12]

def smallest_product_card(img):
    chosen=None
    for parent in img.parents:
        if getattr(parent,"name",None) in ("body","html"): break
        imgs=parent.find_all("img",alt=re.compile(r"^Product\s+",re.I))
        if len(imgs)==1:
            txt=compact(parent.get_text(" ",strip=True))
            if len(txt)>20: chosen=parent
        elif len(imgs)>1:
            break
    return chosen or img.parent

def classify(name,desc):
    t=(" "+name+" "+desc+" ").lower()
    if re.search(r"\b(transceiver|active optical cable|\baoc\b|cfp\d*|qsfp\d*|sfp\+?|optical module)",t): return "Optical Transceivers and AOC"
    if re.search(r"fiber optic|optical fiber|fibre optic|\bmpo\b|\bmtp\b|\bmt ferrule|\blc connector|\bsc connector|expanded beam",t): return "Fiber Optic Connectivity"
    if re.search(r"power distribution|power shelf|power panel|\bpdu\b|distribution panel",t): return "Power Distribution Panels"
    if re.search(r"busbar|bus bar|power connector|barklip|energyedge|power edge|high power|current density|battery connector",t): return "Power Connectors and Busbar"
    if re.search(r"backplane|orthogonal|airmax|examax|metral|hard metric",t): return "Backplane and Orthogonal Connectors"
    if re.search(r"cable assembl|direct attach|\bdac\b|twinax|high.?speed cable|cable connector",t): return "High-Speed Cable Assemblies"
    if re.search(r"board.?to.?board|mezzanine|floating board|fine pitch|micro board|stack height",t): return "High-Speed Board Connectors"
    if re.search(r"dimm|ddr[345]|memory connector|card edge|edge card",t): return "Memory and Card Edge Connectors"
    if re.search(r"pcie|pci express|nvme|sas\b|sata\b|storage connector|u\.2|u\.3|edsff",t): return "Storage and PCIe-SAS Interconnects"
    if re.search(r"ethernet|rj45|usb|hdmi|displayport|external i.?o|input.?output|modular jack",t): return "Ethernet USB and External I-O"
    if re.search(r"ffc|fpc|flex connector|wire.?to.?board|board.?to.?wire|wire.?to.?wire",t): return "Wire-to-Board and FFC-FPC"
    if re.search(r"\brf\b|coax|sma\b|smp\b|smpm|bnc\b|n.?type|mcx|mmcx|ssma|tnc\b",t): return "RF and Coaxial Connectivity"
    if re.search(r"d.?sub|circular|mil.?d|rugged|ip6[5-9]|bayonet connector",t): return "Rugged Circular and D-Sub"
    if re.search(r"sensor|thermistor|temperature|pressure|humidity|material|vent|antenna",t): return "Sensors Materials and Other"
    return "Terminal Blocks and General Interconnect"

def key_spec(desc):
    pats=[
      r"up to\s+\d+(?:\.\d+)?\s*(?:gb/?s|gbps|gt/?s|a|w|kw|ghz)",
      r"\d+(?:\.\d+)?\s*(?:gb/?s|gbps|gt/?s|ghz)",
      r"\d+\s*(?:a|w|kw)\s*(?:per\s+contact)?",
      r"\d+(?:/\d+)?\s*(?:fiber|fibre|position|contact)s?",
      r"\d+(?:\.\d+)?\s*mm\s*pitch",
    ]
    low=desc.lower()
    vals=[]
    for p in pats:
        for m in re.finditer(p,low,re.I):
            v=compact(m.group(0))
            if v not in vals: vals.append(v)
            if len(vals)>=3: return " · ".join(vals)
    return "See official IT Datacom description"

def application(desc):
    t=desc.lower(); a=[]
    rules=[("Data center",r"data center|datacenter|cloud"),("AI/HPC",r"\bai\b|hpc|high.performance computing"),
           ("Server",r"server"),("Storage",r"storage|nvme|sas|sata"),("Networking",r"network|ethernet|infiniband"),
           ("Telecom",r"telecom|wireless|metro"),("Industrial",r"industrial|robot|factory")]
    for label,p in rules:
        if re.search(p,t): a.append(label)
    return " / ".join(a[:4]) if a else "IT Datacom"

def scrape_page(page):
    url=f"{BASE}?pagenumber={page}"
    r=session.get(url,timeout=45)
    r.raise_for_status()
    soup=BeautifulSoup(r.text,"html.parser")
    count_text=compact(soup.get_text(" ",strip=True))
    m=re.search(r"Showing\s+\d+\s+to\s+\d+\s+of\s+(\d+)\s+items",count_text,re.I)
    if m and int(m.group(1))!=EXPECTED:
        raise RuntimeError(f"Official source total changed: {m.group(1)}")
    products=[]
    seen_names=set()
    for img in soup.find_all("img",alt=re.compile(r"^Product\s+",re.I)):
        name=compact(re.sub(r"^Product\s+","",img.get("alt",""),flags=re.I))
        if not name or name in seen_names: continue
        card=smallest_product_card(img)
        strings=[compact(x) for x in card.stripped_strings]
        anchors=card.find_all("a",href=True)
        business=""
        business_url=""
        for a in anchors:
            txt=compact(a.get_text(" ",strip=True))
            href=urljoin(url,a.get("href"))
            if txt and txt!=name and "amphenol.com/markets/it-datacom" not in href:
                business=txt; business_url=href
        clean=[]
        for x in strings:
            if not x or x==name or x==business or x.lower().startswith("image: product"): continue
            if x not in clean: clean.append(x)
        desc=compact(" ".join(clean))
        if desc.startswith(name): desc=compact(desc[len(name):])
        desc=re.sub(r"\s*Showing\s+\d+\s+to\s+\d+\s+of\s+\d+\s+items.*$","",desc,flags=re.I)
        if len(desc)>1600: desc=desc[:1600].rsplit(" ",1)[0]+"…"
        products.append({"name":name,"description":desc,"business":business or "Amphenol","businessUrl":business_url,"sourcePage":page,"officialUrl":url})
        seen_names.add(name)
    return products

def main():
    rows=[]
    for p in range(1,35):
        batch=scrape_page(p)
        print(f"page {p}: {len(batch)}")
        rows.extend(batch)
        time.sleep(.25)
    # Exact de-duplication by name; official listing is expected to have 665 unique visible products.
    dedup={}
    for x in rows:
        if x["name"] not in dedup: dedup[x["name"]]=x
        else:
            # Keep a deterministic suffixed key later if the official listing contains same title with differing text.
            if x["description"]!=dedup[x["name"]]["description"]:
                dedup[x["name"]+" [p"+str(x["sourcePage"])+"]"]=x
    rows=list(dedup.values())
    if len(rows)!=EXPECTED:
        raise RuntimeError(f"Expected {EXPECTED} unique products, parsed {len(rows)}. Refusing partial catalog update.")

    grouped={k:[] for k in CATEGORIES}
    for x in rows:
        x["category"]=classify(x["name"],x["description"])
        grouped[x["category"]].append(x)

    ROOT.mkdir(parents=True,exist_ok=True)
    all_products={}
    used=set()
    def add_product(target,x):
        base=slug(x["name"]); k=base; n=2
        while k in used:
            k=f"{base}-{n}"; n+=1
        used.add(k)
        target[k]={
          "name":x["name"],
          "description":x["description"],
          "officialUrl":x["officialUrl"],
          "businessUrl":x["businessUrl"],
          "specs":{
            "Product Type":x["category"],
            "Key Spec":key_spec(x["description"]),
            "Application":application(x["description"]),
            "Amphenol Business":x["business"],
            "Source Page":f"IT Datacom p.{x['sourcePage']}"
          }
        }

    # Master: all 665 products.
    used=set()
    for x in rows:add_product(all_products,x)
    master_dir=ROOT/"All IT Datacom Products"; master_dir.mkdir(exist_ok=True)
    master={
      "schemaVersion":1,"company":"Amphenol","category":"All IT Datacom Products","displayName":"All IT Datacom Products",
      "description":"All 665 products currently shown on the official Amphenol IT Datacom market listing. Product text is source-derived; DataCenter Tool category labels are tool-side classification.",
      "checked":CHECKED,"officialUrl":BASE,"sourceCount":len(rows),"sourceTotal":EXPECTED,
      "storageMode":"metadata only; official listing and Amphenol business URLs retained",
      "comparisonFields":["Product Type","Key Spec","Application","Amphenol Business","Source Page"],"products":all_products}
    (master_dir/"catalog.json").write_text(json.dumps(master,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
    (master_dir/"readme.md").write_text(f"# All IT Datacom Products\n\nOfficial source: {BASE}\n\nProducts: {len(rows)} / {EXPECTED}\n\nUpdated: {CHECKED}\n",encoding="utf-8")

    # Classified views.
    for cat,items in grouped.items():
        d=ROOT/cat; d.mkdir(parents=True,exist_ok=True)
        products={}; used=set()
        for x in items:add_product(products,x)
        pages=sorted({x["sourcePage"] for x in items})
        manifest={
          "schemaVersion":1,"company":"Amphenol","category":cat,"displayName":cat,
          "description":f"Amphenol IT Datacom official products classified for the DataCenter Tool: {cat}. Product names/descriptions come from the official listing.",
          "checked":CHECKED,"officialUrl":BASE,
          "sourceUrls":[f"{BASE}?pagenumber={p}" for p in pages],
          "sourceCount":len(items),"sourceTotal":EXPECTED,
          "storageMode":"metadata only; official listing and Amphenol business URLs retained",
          "comparisonFields":["Product Type","Key Spec","Application","Amphenol Business","Source Page"],
          "products":products}
        (d/"catalog.json").write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
        (d/"readme.md").write_text(f"# {cat}\n\nOfficial source: {BASE}\n\nClassified products: {len(items)}\nOfficial IT Datacom total: {EXPECTED}\nUpdated: {CHECKED}\n",encoding="utf-8")

    summary={"checked":CHECKED,"officialUrl":BASE,"sourceTotal":EXPECTED,"parsedTotal":len(rows),
             "categoryCounts":{k:len(v) for k,v in grouped.items()}}
    (ROOT/"catalog-summary.json").write_text(json.dumps(summary,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
    print(json.dumps(summary,ensure_ascii=False,indent=2))

if __name__=="__main__":
    main()
