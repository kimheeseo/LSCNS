#!/usr/bin/env python3
from __future__ import annotations
import json, os, re, shutil, sys, zipfile, tempfile, subprocess, urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path
from urllib.parse import quote

ROOT = Path(__file__).resolve().parents[1]
CAT = ROOT / "product_catalog"
COM = CAT / "Commscope"
COR = CAT / "Corning"
CHECKED = "2026-10-05"

ODF_URL = "https://www.commscope.com/network-type/data-centers/optical-distribution-frames-odf/"
PROPEL_URL = "https://www.commscope.com/network-type/central-officeheadend/fiber-panels-and-modules/propel/propel-panels/"
FG_URL = "https://www.commscope.com/network-type/central-officeheadend/fiber-cable-management/fiberguide/"
FA_URL = "https://www.commscope.com/network-type/data-centers/fiber-cable-assemblies/"
CORNING_URL = "https://ecatalog.corning.com/optical-communications/US/en/"

def txt(v): return "" if v is None else str(v).strip()
def write_json(path,obj):
    path.parent.mkdir(parents=True,exist_ok=True)
    path.write_text(json.dumps(obj,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
def github_blob_url(rel):
    return "https://github.com/kimheeseo/LSCNS/blob/main/DCI/DataCenter/product_catalog/"+"/".join(quote(p) for p in str(rel).replace("\\\\","/").split("/"))

def xlsx_rows(path,sheet_name="Products"):
    NS={"m":"http://schemas.openxmlformats.org/spreadsheetml/2006/main","r":"http://schemas.openxmlformats.org/officeDocument/2006/relationships","pr":"http://schemas.openxmlformats.org/package/2006/relationships"}
    with zipfile.ZipFile(path) as z:
        shared=[]
        if "xl/sharedStrings.xml" in z.namelist():
            root=ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall("m:si",NS):
                shared.append("".join(t.text or "" for t in si.iterfind(".//m:t",NS)))
        wb=ET.fromstring(z.read("xl/workbook.xml")); rels=ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        relmap={r.attrib["Id"]:r.attrib["Target"] for r in rels.findall("pr:Relationship",NS)}
        target=None
        for sh in wb.findall("m:sheets/m:sheet",NS):
            if sh.attrib.get("name")==sheet_name:
                target=relmap[sh.attrib["{"+NS["r"]+"}id"]]; break
        if not target:
            sh=wb.find("m:sheets/m:sheet",NS); target=relmap[sh.attrib["{"+NS["r"]+"}id"]]
        if target.startswith("/"): target=target[1:]
        elif not target.startswith("xl/"): target="xl/"+target
        ws=ET.fromstring(z.read(target)); out=[]
        for row in ws.findall(".//m:sheetData/m:row",NS):
            vals={}; maxc=-1
            for c in row.findall("m:c",NS):
                ref=c.attrib.get("r","A1"); letters=re.match(r"[A-Z]+",ref).group(0); col=0
                for ch in letters: col=col*26+(ord(ch)-64)
                col-=1; maxc=max(maxc,col); typ=c.attrib.get("t"); v=c.find("m:v",NS)
                if typ=="s" and v is not None: val=shared[int(v.text)]
                elif typ=="inlineStr": val="".join(t.text or "" for t in c.iterfind(".//m:t",NS))
                elif v is not None: val=v.text
                else: val=""
                vals[col]=val
            if maxc>=0: out.append([vals.get(i,"") for i in range(maxc+1)])
        if not out: return []
        hdr=[txt(x) for x in out[0]]; rows=[]
        for r in out[1:]:
            r=r+[""]*(len(hdr)-len(r))
            if any(txt(x) for x in r): rows.append({hdr[i]:txt(r[i]) for i in range(len(hdr))})
        return rows

def first_match(patterns,s,flags=re.I):
    for p in patterns:
        m=re.search(p,s,flags)
        if m: return m.group(1) if m.lastindex else m.group(0)
    return ""
def parse_fiber_count(desc,ptype=""):
    m=re.search(r"(\d+)\s*x\s*(\d+)\s*fiber\s*MPO",desc,re.I)
    if m:return str(int(m.group(1))*int(m.group(2)))
    for p in [r"(\d+)\s*[- ]?\s*Fiber\b",r"(\d+)\s*fibers?\b",r"(\d+)\s*color kit"]:
        m=re.search(p,desc,re.I)
        if m:return m.group(1)
    if "duplex" in ptype.lower():return "2"
    if "simplex" in ptype.lower():return "1"
    return ""
def parse_fiber_mode(desc):
    for k in ("OM5","OM4","OM3","OM2","OM1","OS2"):
        if re.search(r"\b"+k+r"\b",desc,re.I):return k
    if re.search(r"single.?mode",desc,re.I):return "Singlemode"
    if re.search(r"multi.?mode",desc,re.I):return "Multimode"
    return ""
def parse_connectors(desc):
    pat=re.compile(r"(MPO(?:-?\d+)?(?:/(?:APC|UPC))?(?:\s*\((?:Non-Pinned|Pinned)\))?|MTP(?:-?\d+)?(?:/(?:APC|UPC))?|LC/(?:UPC|APC)|SC/(?:UPC|APC)(?:\s*9°)?|LSH/(?:UPC|APC)|MDC(?:/(?:UPC|APC))?|MMC(?:/(?:UPC|APC))?|SN(?:/(?:UPC|APC))?)",re.I)
    vals=[]
    for m in pat.finditer(desc):
        v=re.sub(r"\s*\((?:Non-Pinned|Pinned)\)","",m.group(1),flags=re.I)
        v=re.sub(r"\b(MPO|MTP)(\d+)\b",r"\1-\2",v,flags=re.I)
        v=re.sub(r"\s+"," ",v).strip()
        if v.lower() not in [x.lower() for x in vals]: vals.append(v)
    return " / ".join(vals[:2])
def parse_gender(desc):
    if re.search(r"Non-Pinned",desc,re.I):return "Female"
    if re.search(r"\bPinned\b",desc,re.I):return "Male"
    return ""
def parse_jacket(desc):
    v=first_match([r"\b(LZSH\s*B2ca)\b",r"\b(LSZH(?:/OFNR)?)\b",r"\b(OFNP)\b",r"\b(OFNR)\b",r"\b(Plenum)\b",r"\b(Riser)\b",r"\b(Low Smoke Zero Halogen)\b"],desc)
    return re.sub(r"^LZSH", "LSZH", v, flags=re.I) if v else ""
def parse_length(desc):
    ms=list(re.finditer(r"(\d+(?:\.\d+)?)\s*(m|ft)\b",desc,re.I))
    return (ms[-1].group(1)+" "+ms[-1].group(2)) if ms else ""
def parse_color(desc): return first_match([r"\b(Yellow|Black|Gray|Grey|Aqua|White|Lime green)\b"],desc)
def parse_dimensions(desc):
    vals=[]
    for m in re.finditer(r"(\d+(?:\.\d+)?\s*x\s*\d+(?:\.\d+)?\s*in)",desc,re.I):
        v=re.sub(r"\s+","",m.group(1))
        if v not in vals:vals.append(v)
    return " | ".join(vals[:5])
def fa_family(ptype):
    t=ptype.lower()
    if "trunk" in t:return "Trunk Assemblies","trunk-assemblies"
    if "patch cord" in t:return "Patch Cords","patch-cords"
    if "pigtail" in t:return "Pigtails","pigtails"
    return "Arrays and Fanouts","arrays-and-fanouts"
def product_url_fa(part,segment):
    p=part.lower()
    return f"https://www.commscope.com/network-type/data-centers/fiber-cable-assemblies/{segment}/{p}/" if re.fullmatch(r"[a-z0-9._-]+",p) else FA_URL
def product_url_fg(part):
    p=part.lower()
    return "https://www.commscope.com/product-type/cable-management/raceways/item"+p+"/" if re.fullmatch(r"[a-z0-9._-]+",p) else FG_URL

def build_commscope():
    odf=xlsx_rows(COM/"ODF.xlsx"); fg=xlsx_rows(COM/"FiberGuide® Fiber Raceways for Central OfficeHeadend.xlsx"); fa=xlsx_rows(COM/"Fiber Cable Assemblies.xlsx")
    for d in [COM/"FiberGuide",COM/"Fiber Cable Assemblies",COM/"Propel Panels"]:
        if d.exists():shutil.rmtree(d)
    pdfs=list((COM/"ODF").glob("*.pdf")); pdf_map={}
    for p in pdfs:
        key=re.sub(r"^(CS_|P360_)","",p.stem,flags=re.I); key=re.sub(r"_external$","",key,flags=re.I); pdf_map[key.upper()]=p

    odf_products={}; propel_rows=[]
    for r in odf:
        if r.get("Product Brand")=="Propel" or r.get("Product Series")=="PPL" or "PROPEL" in (r.get("Part Number","")+" "+r.get("Part Name","")).upper():
            propel_rows.append(r);continue
        part=r.get("Part Number") or r.get("Part Name")
        if not part:continue
        desc=r.get("Description","");ptype=r.get("Product Type","");family=re.sub(r"^Variant of\s+","",r.get("Part Indicator",""),flags=re.I).strip().upper();src=pdf_map.get(family)
        fc=parse_fiber_count(desc,ptype);fm=parse_fiber_mode(desc);co=parse_connectors(desc);gender=parse_gender(desc)
        specs={"Part Number":part,"Part Name":r.get("Part Name",""),"Product Type":ptype,"Product Brand":r.get("Product Brand",""),"Product Series":r.get("Product Series",""),
               "Fiber Count":fc or "—","Fiber Mode":fm or "—","Fiber Type":fm or "—","Connector Type":co or "—","Connector":co or "—","Gender":gender or "—",
               "Port Count":first_match([r"(\d+)-port"],desc) or "—","Source Family":family or "—"}
        if src:specs["Source PDF"]=src.name
        odf_products[part]={"name":" · ".join(x for x in [part,r.get("Part Name"),ptype] if x),"description":desc,
                            "officialUrl":github_blob_url(src.relative_to(CAT)) if src else ODF_URL,"businessUrl":ODF_URL,"checked":CHECKED,"specs":specs}
    write_json(COM/"ODF"/"catalog.json",{"schemaVersion":1,"company":"CommScope","category":"Optical Distribution Frames / Panels","displayName":"CommScope ODF · FACT® / NG4access® / FIST®",
      "description":"CommScope data-center ODF products generated from the uploaded ODF product list and linked family PDFs.","checked":CHECKED,"officialUrl":ODF_URL,"sourceUrls":[ODF_URL],
      "sourceFiles":["ODF.xlsx"]+[p.name for p in pdfs],"comparisonFields":["Product Type","Product Brand","Product Series","Fiber Count","Fiber Mode","Connector Type","Port Count","Source Family"],"products":odf_products})

    products={}
    static=[
      ("760252002","PPL-1U","1","Sliding","72","72","144","288","Gray","https://www.commscope.com/network-type/data-centers/fiber-panels-and-modules/propel/item760252002/"),
      ("760252003","PPL-2U","2","Sliding","144","144","288","576","Gray","https://www.commscope.com/network-type/data-centers/fiber-panels-and-modules/propel/item760252003/"),
      ("760252004","PPL-4U","4","Sliding","288","288","576","1152","Gray","https://www.commscope.com/network-type/data-centers/fiber-panels-and-modules/propel/item760252004/"),
      ("760255953","PPL-1U-W","1","Sliding","72","72","144","288","White",PROPEL_URL),("760255954","PPL-2U-W","2","Sliding","144","144","288","576","White",PROPEL_URL),
      ("760255955","PPL-4U-W","4","Sliding","288","288","576","1152","White",PROPEL_URL),
      ("760253796","PPL-1U-HD-FX","1","Fixed","48","48","96","192","Gray","https://www.commscope.com/network-type/central-officeheadend/fiber-panels-and-modules/propel/item760253796/"),
      ("760253797","PPL-2U-HD-FX","2","Fixed","96","96","192","384","Gray","https://www.commscope.com/network-type/data-centers/fiber-panels-and-modules/propel-ppl-panels/item760253797/"),
      ("760259115","PPL-1U-ANG-FX","1","Fixed","72","72","144","288","Gray",PROPEL_URL)]
    for part,name,ru,movement,lc,mpo,sn,fibers,color,url in static:
        specs={"Part Number":part,"Product Type":"Fiber patch panel","Product Brand":"Propel","Product Series":"PPL","Rack Units":ru,"Shelf Movement":movement,
               "Max Duplex LC Ports":lc,"Max MPO Ports":mpo,"Max SN Ports":sn,"Maximum Fiber Count":fibers,"Color":color,"Supported Fiber Counts":"8 / 12 / 16 / 24",
               "Labeling":"1 lane = 2 MPO ports or 2 duplex LC pairs"}
        products[part]={"name":part+" · "+name+" · Fiber patch panel","description":f"Propel {ru}RU {movement.lower()} fiber panel; up to {lc} duplex LC, {mpo} MPO or {sn} SN ports ({fibers}f).","officialUrl":url,"businessUrl":url,"checked":CHECKED,"specs":specs}
    for r in propel_rows:
        part=r.get("Part Number") or r.get("Part Name")
        if not part or part in products:continue
        desc=r.get("Description","");ptype=r.get("Product Type","") or "Propel XFrame"
        products[part]={"name":" · ".join(x for x in [part,r.get("Part Name"),ptype] if x),"description":desc,"officialUrl":PROPEL_URL,"businessUrl":PROPEL_URL,"checked":CHECKED,
                        "specs":{"Part Number":part,"Part Name":r.get("Part Name",""),"Product Type":ptype,"Product Brand":"Propel","Product Series":"PPL",
                                 "Rack Units":first_match([r"(\d+)RU"],desc) or "—","Supported Fiber Counts":"8 / 12 / 16 / 24 (model dependent)",
                                 "Labeling":"1 lane = 2 MPO ports or 2 duplex LC pairs"}}
    prop_count=len(products)
    write_json(COM/"Propel Panels"/"catalog.json",{"schemaVersion":1,"company":"CommScope","category":"Fiber Patch Panel","displayName":"Propel® Panels / XFrame",
      "description":"Propel panel portfolio including PPL sliding/fixed panels and XFrame products. Labeling rule is derived from the uploaded PPL labeling template.","checked":CHECKED,
      "officialUrl":PROPEL_URL,"sourceUrls":[PROPEL_URL],"sourceFiles":["Propel-1U-2U-4U-Labeling-Template.xlsx","CS_PROPEL-PANELS_external.pdf"],
      "comparisonFields":["Rack Units","Shelf Movement","Max Duplex LC Ports","Max MPO Ports","Max SN Ports","Maximum Fiber Count","Supported Fiber Counts","Color"],"products":products})

    products={}
    for r in fg:
        part=r.get("Part Number") or r.get("Part Name")
        if not part:continue
        desc=r.get("Description","");ptype=r.get("Product Type","");url=product_url_fg(part)
        specs={"Part Number":part,"Product Type":ptype,"Product Brand":r.get("Product Brand") or "FiberGuide®","System Dimensions":parse_dimensions(desc) or "—",
               "Length":parse_length(desc) or "—","Color":parse_color(desc) or "—","Section / Function":ptype or "—"}
        products[part]={"name":" · ".join(x for x in [part,ptype] if x),"description":desc,"officialUrl":url,"businessUrl":url,"checked":CHECKED,"specs":specs}
    write_json(COM/"FiberGuide"/"catalog.json",{"schemaVersion":1,"company":"CommScope","category":"Cable Management FiberGuide","displayName":"FiberGuide® Fiber Raceway",
      "description":"Complete FiberGuide product list generated from the uploaded CommScope Excel product export.","checked":CHECKED,"officialUrl":FG_URL,"sourceUrls":[FG_URL],
      "sourceFiles":["FiberGuide® Fiber Raceways for Central OfficeHeadend.xlsx"],"comparisonFields":["Product Type","System Dimensions","Length","Color","Section / Function"],"products":products})

    groups={}
    for r in fa:
        family,segment=fa_family(r.get("Product Type",""));groups.setdefault((family,segment),[]).append(r)
    for (family,segment),rows in groups.items():
        for chunk_idx in range(0,len(rows),400):
            chunk=rows[chunk_idx:chunk_idx+400];products={}
            for r in chunk:
                part=r.get("Part Number") or r.get("Part Name")
                if not part:continue
                desc=r.get("Description","");ptype=r.get("Product Type","");fc=parse_fiber_count(desc,ptype);fm=parse_fiber_mode(desc);co=parse_connectors(desc);gender=parse_gender(desc);jacket=parse_jacket(desc);ln=parse_length(desc);url=product_url_fa(part,segment)
                specs={"Part Number":part,"Part Name":r.get("Part Name",""),"Product Type":ptype,"Product Brand":r.get("Product Brand",""),"Product Series":r.get("Product Series",""),
                       "Fiber Count":fc or "—","Fiber Mode":fm or "—","Fiber Type":fm or "—","Connector Type":co or "—","Connector":co or "—","Gender":gender or "—",
                       "Cable Jacket":jacket or "—","Jacket":jacket or "—","Length":ln or "—"}
                products[part]={"name":" · ".join(x for x in [part,r.get("Part Name"),ptype] if x),"description":desc,"officialUrl":url,"businessUrl":url,"checked":CHECKED,"specs":specs}
            n=chunk_idx//400+1
            write_json(COM/"Fiber Cable Assemblies"/family/f"Part {n:02d}"/"catalog.json",{"schemaVersion":1,"company":"CommScope","category":"Fiber Cable Assemblies - "+family,
              "displayName":family+f" · Part {n:02d}","description":"CommScope fiber cable assembly SKUs generated from the uploaded product-list workbook.","checked":CHECKED,"officialUrl":FA_URL,
              "sourceUrls":[FA_URL],"sourceFiles":["Fiber Cable Assemblies.xlsx"],"comparisonFields":["Product Type","Product Brand","Product Series","Fiber Count","Fiber Mode","Connector Type","Gender","Cable Jacket","Length"],"products":products})

    (COM/"readme.md").write_text(f"# CommScope Data Center Product Catalog\n\nGenerated from the uploaded CommScope product-list workbooks and official CommScope product pages.\n\n- ODF / FACT / NG4access: {len(odf_products)} structured products\n- Propel / XFrame: {prop_count} structured panel/frame entries\n- FiberGuide: {len(fg)} products\n- Fiber Cable Assemblies: {len(fa)} products\n- Cable fields: Fiber Count / Fiber Mode / Connector Type / Cable Jacket\n- Source PDFs under ODF remain unchanged and are linked by family where applicable.\n- PPL labeling template rule: one lane corresponds to two MPO ports or two duplex LC pairs.\n\nOfficial sources:\n- {PROPEL_URL}\n- {ODF_URL}\n- {FG_URL}\n- {FA_URL}\n",encoding="utf-8")

def corning_mode(model):
    m=re.match(r"^(\d{3})([A-Z0-9]{3})-",model)
    if not m:return "—"
    code=m.group(2);suffix=model.split("-",1)[1] if "-" in model else ""
    if code.startswith(("E","Z","D","F","P")):return "OS2"
    if code.startswith("K"):return "OM1"
    if code.startswith("T"):
        if "90" in suffix:return "OM4"
        if "80" in suffix:return "OM3"
        return "OM2"
    if code.startswith("J"):return "Multimode"
    return "—"
def corning_assembly_mode(model):
    if "89G" in model:return "OS2"
    if "93Q" in model:return "OM4"
    if "93T" in model:return "OM3"
    m=re.search(r"(?:\d{2}|E4)([QTG])(?=[A-Z])",model)
    return {"Q":"OM4","T":"OM3","G":"OS2"}.get(m.group(1),"—") if m else "—"
def corning_count(model,top):
    if top=="Module":
        m=re.search(r"(?:UM|RM|MM|SM)(\d{2})",model)
        if m:return str(int(m.group(1)))
    m=re.search(r"(\d{2})([QTG])(?=[A-Z])",model)
    if m and int(m.group(1)) in (4,8,12,16,24,32,36,48,72,96):return str(int(m.group(1)))
    if re.search(r"E4[QTG]",model):return "144"
    m=re.match(r"^(\d{3})",model);return str(int(m.group(1))) if m else "—"

def build_corning_missing():
    for folder,dirs,files in os.walk(COR):
        p=Path(folder);pdfs=sorted([x for x in files if x.lower().endswith(".pdf")])
        if not pdfs or (p/"catalog.json").exists():continue
        rel=p.relative_to(COR);top=rel.parts[0] if rel.parts else "Other";family=rel.parts[-1];products={};cfields=[];category=top
        for fn in pdfs:
            model=re.sub(r"_NAFTA_AEN$|_AEN$","",Path(fn).stem,flags=re.I)
            if top=="Cable":
                mode=corning_mode(model);specs={"Fiber Count":corning_count(model,top),"Fiber Mode":mode,"Fiber Type":mode,"Connector Type":"Unterminated","Connector":"Unterminated",
                  "Cable Jacket":"Polyethylene (PE)","Jacket":"Polyethylene (PE)","Cable Construction":family,"Armoring":"Yes" if re.search(r"Armored|Lite",family,re.I) and "All-Dielectric" not in family else ("No" if "All-Dielectric" in family else "—")}
                cfields=["Fiber Count","Fiber Mode","Connector Type","Cable Jacket","Cable Construction","Armoring"];category="Fiber Cable"
            elif top in ("Harness","Trunk","Module"):
                mode=corning_assembly_mode(model);count=corning_count(model,top)
                if top=="Harness":connector="MTP / LC" if re.search(r"MTP to LC|Staggered|Conversion",family,re.I) else ("MTP / MTP" if re.search(r"Y Harness",family,re.I) else "See PDF");category="Fiber Harness"
                elif top=="Trunk":connector="MTP / MTP" if "MTP" in family else "See PDF";category="Fiber Trunk"
                else:connector="LC Duplex / MTP" if re.search(r"Base-8|Ultra Low Loss|Fiber to the Desk",family,re.I) else "See PDF";category="Fiber Module"
                specs={"Fiber Count":count,"Fiber Mode":mode,"Fiber Type":mode,"Connector Type":connector,"Connector":connector}
                if top in ("Harness","Trunk"):specs.update({"Cable Jacket":"—","Jacket":"—","Length":first_match([r"(\d{3}[FM])$"],model) or "—"});cfields=["Fiber Count","Fiber Mode","Connector Type","Cable Jacket","Length"]
                else:cfields=["Fiber Count","Fiber Mode","Connector Type"]
            else:
                category={"Panel":"Fiber Patch Panel","ODF":"Fiber Patch Panel ODF","Connector":"Fiber Connector","Accessories":"Fiber Component Accessories"}.get(top,top)
                connector="MTP" if "MTP" in family else ("MMC" if "MMC" in family else ("FC" if "FC" in family else ("SC" if "SC" in family else "—")));specs={"Product Type":family,"Connector Type":connector,"Connector":connector};cfields=["Product Type","Connector Type"]
            pdf_path=p/fn;url=github_blob_url(pdf_path.relative_to(CAT))
            products[model]={"name":model+" · "+family,"description":family,"officialUrl":url,"businessUrl":url,"checked":CHECKED,"specs":specs}
        write_json(p/"catalog.json",{"schemaVersion":1,"company":"Corning","category":category,"displayName":family,"description":"Structured metadata generated from uploaded Corning product specification PDFs.","checked":CHECKED,
          "officialUrl":CORNING_URL,"sourceType":"Uploaded Corning product specification PDFs","pdfCount":len(pdfs),"comparisonFields":cfields,"products":products})


# Current CommScope product-list exports. These are the same Excel exports exposed by
# the official product/category pages, so the catalog can be refreshed without
# manually committing a new binary workbook each time.
COMMSCOPE_EXPORTS={
 "Cable Management":("https://www.commscope.com/product-type/cable-management/download/?id=1073742080","https://www.commscope.com/product-type/cable-management/"),
 "Fiber Optic Cables":("https://www.commscope.com/product-type/cables/fiber-cables/download/?id=1073742104","https://www.commscope.com/product-type/cables/fiber-cables/"),
 "Fiber Cables Data Center":("https://www.commscope.com/network-type/data-centers/fiber-cables/download/?id=1073742311","https://www.commscope.com/network-type/data-centers/fiber-cables/"),
 "Building Entrance Solutions":("https://www.commscope.com/network-type/data-centers/building-entrance-solutions/download/?id=1073742322","https://www.commscope.com/network-type/data-centers/building-entrance-solutions/"),
 "Fiber Panels Modules Cassettes":("https://www.commscope.com/product-type/frames-panels-cassettes-modules/fiber-panels-modules-cassettes/download/?id=1073742159","https://www.commscope.com/product-type/frames-panels-cassettes-modules/fiber-panels-modules-cassettes/"),
 "Twisted Pair Cable Assemblies":("https://www.commscope.com/product-type/cable-assemblies/twisted-pair-cable-assemblies/download/?id=1073742068","https://www.commscope.com/product-type/cable-assemblies/twisted-pair-cable-assemblies/"),
 "Coaxial Cables":("https://www.commscope.com/product-type/cables/coaxial-cables/download/?id=1073742101","https://www.commscope.com/product-type/cables/coaxial-cables/"),
 "Twisted Pair Cables":("https://www.commscope.com/product-type/cables/twisted-pair-cables/download/?id=1073742111","https://www.commscope.com/product-type/cables/twisted-pair-cables/"),
 "Cable Assemblies":("https://www.commscope.com/product-type/cable-assemblies/download/?id=1073742060","https://www.commscope.com/product-type/cable-assemblies/"),
 "Copper Panels Modules Cassettes":("https://www.commscope.com/product-type/frames-panels-cassettes-modules/copper-panels-modules-cassettes/download/?id=1073742170","https://www.commscope.com/product-type/frames-panels-cassettes-modules/copper-panels-modules-cassettes/")
}

def download_xlsx(url,dest):
    ua="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/153 Safari/537.36"
    req=urllib.request.Request(url,headers={"User-Agent":ua,"Accept":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*","Referer":"https://www.commscope.com/"})
    try:
        with urllib.request.urlopen(req,timeout=45) as r:
            data=r.read()
        if data[:2]==b"PK" and len(data)>2000:
            Path(dest).write_bytes(data);return True
    except Exception as e:
        print("urllib download failed",url,e)
    try:
        subprocess.run(["curl","-L","--fail","--retry","2","--max-time","60","-A",ua,"-H","Referer: https://www.commscope.com/","-o",str(dest),url],check=True)
        data=Path(dest).read_bytes()
        if data[:2]==b"PK" and len(data)>2000:return True
    except Exception as e:
        print("curl download failed",url,e)
    if Path(dest).exists():Path(dest).unlink()
    return False

def current_export_rows(tmp,name):
    url,_=COMMSCOPE_EXPORTS[name];dest=Path(tmp)/(re.sub(r"[^A-Za-z0-9]+","_",name)+".xlsx")
    if not download_xlsx(url,dest):
        print("WARN: no current export for",name);return []
    rows=xlsx_rows(dest)
    print("downloaded",name,len(rows))
    return rows

def catalog_part_numbers(folder):
    ids=set()
    if not folder.exists():return ids
    for p in folder.rglob("catalog.json"):
        try:
            data=json.loads(p.read_text(encoding="utf-8"))
            for key,meta in (data.get("products") or {}).items():
                ids.add(txt((meta.get("specs") or {}).get("Part Number")) or txt(key))
        except Exception as e:print("catalog read warning",p,e)
    return ids

def row_part(r):return txt(r.get("Part Number")) or txt(r.get("Part Name"))
def dedupe_rows(rows,used):
    out=[];seen=set()
    for r in rows:
        p=row_part(r)
        if not p or p in used or p in seen:continue
        seen.add(p);used.add(p);out.append(r)
    return out
def union_rows(*groups):
    d={}
    for rows in groups:
        for r in rows:
            p=row_part(r)
            if p and p not in d:d[p]=r
    return list(d.values())

def copper_category(desc):
    m=re.search(r"\bCat(?:egory)?\s*(5e|6A|6|7A|7|8)\b",desc,re.I)
    return ("Cat"+m.group(1)) if m else "—"
def pair_count(desc):
    m=re.search(r"\b(\d+)\s*(?:pair|pairs)\b",desc,re.I);return m.group(1) if m else "—"
def awg(desc):
    m=re.search(r"\b(\d+)\s*AWG\b",desc,re.I);return m.group(1)+" AWG" if m else "—"
def shielding(desc):
    m=re.search(r"\b(U/UTP|F/UTP|S/FTP|U/FTP|F/FTP|SF/UTP)\b",desc,re.I);return m.group(1).upper() if m else "—"
def generic_connector(desc):
    c=parse_connectors(desc)
    if c:return c
    if re.search(r"\bRJ45\b",desc,re.I):return "RJ45"
    if re.search(r"\bLSA[- ]?PLUS\b",desc,re.I):return "LSA-PLUS"
    if re.search(r"\bBNC\b",desc,re.I):return "BNC"
    if re.search(r"\bF[- ]?type\b",desc,re.I):return "F-Type"
    return "—"
def port_count(desc):
    m=re.search(r"\b(\d+)[- ]?port\b",desc,re.I);return m.group(1) if m else "—"
def rack_units(desc):
    m=re.search(r"\b(\d+(?:\.\d+)?)\s*(?:RU|U)\b",desc,re.I);return (m.group(1)+"U") if m else "—"
def flame_rating(desc):
    m=re.search(r"\b(B2ca|Bca|Cca|Dca|Eca)\b",desc,re.I);return m.group(1) if m else "—"
def impedance(desc):
    m=re.search(r"\b(\d+)\s*(?:ohm|Ω)\b",desc,re.I);return m.group(1)+" Ohm" if m else "—"
def row_base_specs(r):
    return {"Part Number":row_part(r),"Part Name":txt(r.get("Part Name")),"Product Type":txt(r.get("Product Type")) or "—",
            "Product Brand":txt(r.get("Product Brand")) or "—","Product Series":txt(r.get("Product Series")) or "—",
            "Status":txt(r.get("Status")) or "—"}

def extended_specs(r,kind):
    desc=txt(r.get("Description"));ptype=txt(r.get("Product Type"));s=row_base_specs(r)
    if kind in ("fiber","fiber_panel","entrance"):
        fc=parse_fiber_count(desc,ptype);fm=parse_fiber_mode(desc);co=generic_connector(desc)
        if kind=="fiber" and co=="—":co="Unterminated"
        s.update({"Fiber Count":fc or "—","Fiber Mode":fm or "—","Fiber Type":fm or "—","Connector Type":co,"Connector":co,
                  "Cable Jacket":parse_jacket(desc) or "—","Jacket":parse_jacket(desc) or "—","Flame Rating":flame_rating(desc),
                  "Length":parse_length(desc) or "—"})
        if kind=="fiber_panel":s.update({"Port Count":port_count(desc),"Rack Units":rack_units(desc)})
    elif kind in ("twisted","twisted_assembly","copper_assembly","copper_panel"):
        co=generic_connector(desc);j=parse_jacket(desc) or "—"
        s.update({"Category":copper_category(desc),"Pair Count":pair_count(desc),"AWG":awg(desc),"Shielding":shielding(desc),
                  "Connector Type":co,"Connector":co,"Cable Jacket":j,"Jacket":j,"Flame Rating":flame_rating(desc),
                  "Length":parse_length(desc) or "—"})
        if kind=="copper_panel":s.update({"Port Count":port_count(desc),"Rack Units":rack_units(desc)})
    elif kind=="coax":
        j=parse_jacket(desc) or "—";co=generic_connector(desc)
        s.update({"Impedance":impedance(desc),"Connector Type":co,"Connector":co,"Cable Jacket":j,"Jacket":j,
                  "Flame Rating":flame_rating(desc),"Length":parse_length(desc) or "—"})
    elif kind=="management":
        s.update({"System Dimensions":parse_dimensions(desc) or "—","Length":parse_length(desc) or "—","Color":parse_color(desc) or "—"})
    return s

EXTENDED_CONFIG={
 "Fiber Cables":("Fiber Optic Cable","fiber",["Product Type","Product Brand","Fiber Count","Fiber Mode","Connector Type","Cable Jacket","Flame Rating","Length"]),
 "Fiber Panels Modules Cassettes":("Fiber Panels Modules Cassettes","fiber_panel",["Product Type","Product Brand","Fiber Count","Fiber Mode","Connector Type","Port Count","Rack Units"]),
 "Building Entrance Solutions":("Fiber Building Entrance Solutions","entrance",["Product Type","Product Brand","Fiber Count","Fiber Mode","Connector Type"]),
 "Cable Management":("Cable Management","management",["Product Type","Product Brand","System Dimensions","Length","Color"]),
 "Twisted Pair Cable Assemblies":("Copper Twisted Pair Cable Assemblies","twisted_assembly",["Product Type","Product Brand","Category","Pair Count","AWG","Shielding","Connector Type","Cable Jacket","Length"]),
 "Twisted Pair Cables":("Copper Twisted Pair Cables","twisted",["Product Type","Product Brand","Category","Pair Count","AWG","Shielding","Cable Jacket","Flame Rating","Length"]),
 "Copper Module Cable Assemblies":("Copper Cable Assemblies","copper_assembly",["Product Type","Product Brand","Category","Pair Count","AWG","Shielding","Connector Type","Cable Jacket","Length"]),
 "Copper Panels Modules Cassettes":("Copper Panels Modules Cassettes","copper_panel",["Product Type","Product Brand","Category","Connector Type","Port Count","Rack Units"]),
 "Coaxial Cables":("Copper Coaxial Cables","coax",["Product Type","Product Brand","Impedance","Connector Type","Cable Jacket","Length"])
}

def write_extended_category(name,rows,official_url,source_names):
    category,kind,fields=EXTENDED_CONFIG[name];folder=COM/name
    if folder.exists():shutil.rmtree(folder)
    for i in range(0,len(rows),400):
        chunk=rows[i:i+400];products={}
        for r in chunk:
            part=row_part(r);desc=txt(r.get("Description"));ptype=txt(r.get("Product Type"))
            specs=extended_specs(r,kind)
            products[part]={"name":" · ".join(x for x in [part,txt(r.get("Part Name")),ptype] if x),"description":desc,
                            "officialUrl":official_url,"businessUrl":official_url,"checked":CHECKED,"specs":specs}
        n=i//400+1
        write_json(folder/f"Part {n:02d}"/"catalog.json",{"schemaVersion":1,"company":"CommScope","category":category,
          "displayName":name+f" · Part {n:02d}","description":"CommScope products generated from the current official downloadable Product List Excel.",
          "checked":CHECKED,"officialUrl":official_url,"sourceUrls":[official_url],"sourceFiles":source_names,
          "comparisonFields":fields,"products":products})
    return len(rows)

def build_extended_commscope():
    baseline=set()
    for p in [COM/"ODF",COM/"Propel Panels",COM/"FiberGuide",COM/"Fiber Cable Assemblies"]:
        baseline.update(catalog_part_numbers(p))
    used=set(baseline)
    with tempfile.TemporaryDirectory() as tmp:
        exports={}
        for name in COMMSCOPE_EXPORTS:exports[name]=current_export_rows(tmp,name)
        fiber=union_rows(exports["Fiber Cables Data Center"],exports["Fiber Optic Cables"])
        cable_assemblies=exports["Cable Assemblies"]
        copper_module=[r for r in cable_assemblies if txt(r.get("Product Type")).lower() in ("copper patch cord","copper test cord")]
        ordered=[
          ("Fiber Cables",fiber,COMMSCOPE_EXPORTS["Fiber Optic Cables"][1],["Fiber Cables.xlsx","Fiber Optic Cables.xlsx"]),
          ("Fiber Panels Modules Cassettes",exports["Fiber Panels Modules Cassettes"],COMMSCOPE_EXPORTS["Fiber Panels Modules Cassettes"][1],["Fiber Panels, Modules & Cassettes.xlsx"]),
          ("Building Entrance Solutions",exports["Building Entrance Solutions"],COMMSCOPE_EXPORTS["Building Entrance Solutions"][1],["Building Entrance Solutions.xlsx"]),
          ("Cable Management",exports["Cable Management"],COMMSCOPE_EXPORTS["Cable Management"][1],["Cable Management.xlsx"]),
          ("Twisted Pair Cable Assemblies",exports["Twisted Pair Cable Assemblies"],COMMSCOPE_EXPORTS["Twisted Pair Cable Assemblies"][1],["Twisted Pair Cable Assemblies.xlsx"]),
          ("Twisted Pair Cables",exports["Twisted Pair Cables"],COMMSCOPE_EXPORTS["Twisted Pair Cables"][1],["Twisted Pair Cables.xlsx"]),
          ("Copper Module Cable Assemblies",copper_module,COMMSCOPE_EXPORTS["Cable Assemblies"][1],["Copper Module Cable Assemblies.xlsx"]),
          ("Copper Panels Modules Cassettes",exports["Copper Panels Modules Cassettes"],COMMSCOPE_EXPORTS["Copper Panels Modules Cassettes"][1],["Copper Panels, Modules & Cassettes.xlsx"]),
          ("Coaxial Cables",exports["Coaxial Cables"],COMMSCOPE_EXPORTS["Coaxial Cables"][1],["Coaxial Cables.xlsx"])
        ]
        counts={}
        for name,rows,url,sources in ordered:
            if not rows:
                print("SKIP empty official export",name);continue
            clean=dedupe_rows(rows,used);counts[name]=write_extended_category(name,clean,url,sources)
        print("extended CommScope unique products",sum(counts.values()),counts)
        summary=COM/"readme.md"
        old=summary.read_text(encoding="utf-8") if summary.exists() else "# CommScope Data Center Product Catalog\n"
        old=re.sub(r"\n## Extended current product lists[\s\S]*$","",old).rstrip()
        lines=["","## Extended current product lists",""]
        for name in EXTENDED_CONFIG:
            if name in counts:lines.append(f"- {name}: {counts[name]} unique products")
        lines+=["",f"- Baseline structured products before extended lists: {len(baseline)}",f"- Extended unique products added: {sum(counts.values())}","- Duplicate Part Numbers are retained only once across CommScope catalogs.","- Discontinued products remain searchable in the product list but are excluded from BOM candidate matching.",""]
        summary.write_text(old+"\n"+"\n".join(lines),encoding="utf-8")

def refresh_indexes(revision=""):
    catalogs=sorted(str(p.relative_to(CAT)).replace("\\\\","/") for p in CAT.rglob("catalog.json"))
    manifest=CAT/"catalog-manifest.json";obj={}
    if manifest.exists():
        try:obj=json.loads(manifest.read_text(encoding="utf-8"))
        except:obj={}
    obj["revision"]=revision or obj.get("revision","");obj["catalogTree"]="generated";obj["paths"]=catalogs
    manifest.write_text(json.dumps(obj,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
    idx=CAT/"bom-catalog-index.json"
    if idx.exists():
        try:
            o=json.loads(idx.read_text(encoding="utf-8"));o["checked"]=CHECKED;o["totalStructuredCatalogs"]=len(catalogs);idx.write_text(json.dumps(o,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
        except Exception:pass
    return len(catalogs)

if __name__=="__main__":
    if len(sys.argv)>=3 and sys.argv[1]=="--manifest-only":
        print("manifest catalogs",refresh_indexes(sys.argv[2]));raise SystemExit
    build_commscope();build_extended_commscope();build_corning_missing();print("generated catalogs",len(list(CAT.rglob("catalog.json"))))
