(() => {
'use strict';
const VERSION='7.4.21';
let selectedCategory='all',returnFocus;
const textOf=e=>(e&&(e.innerText||e.textContent)||'').replace(/\s+/g,' ').trim();

const I18N={
  ko:{button:'Supply Chain',title:'BOM Supply Chain · 업체 / 제품 맵',sub:'현재 설계에 실제 사용된 업체·제품·수량을 우선 표시하고, 각 영역별 관련 기업은 별도 참고 목록으로 함께 보여줍니다.',need:'먼저 설계 계산을 실행해 주세요.',center:'현재 BOM',used:'BOM 반영 항목',sources:'제품 / 데이터시트',close:'닫기',refresh:'새로고침',
      cats:{compute:'GPU / Accelerator',cpu:'Server CPU',network:'Network Fabric',optical:'Optical Connectivity',power:'Power / UPS / PDU',cooling:'Cooling / HVAC',rack:'Rack / Physical',storage:'Storage',monitoring:'Fiber Test / Monitoring',facility:'Facility / Security'}},
  en:{button:'Supply Chain',title:'BOM Supply Chain · Vendor / Product Map',sub:'Shows vendors/products actually used by the current BOM first, plus a separate related-companies reference list for each category.',need:'Run the design calculation first.',center:'Current BOM',used:'BOM items',sources:'Product / datasheet',close:'Close',refresh:'Refresh',
      cats:{compute:'GPU / Accelerator',cpu:'Server CPU',network:'Network Fabric',optical:'Optical Connectivity',power:'Power / UPS / PDU',cooling:'Cooling / HVAC',rack:'Rack / Physical',storage:'Storage',monitoring:'Fiber Test / Monitoring',facility:'Facility / Security'}},
  ja:{button:'サプライチェーン',title:'BOM サプライチェーン',sub:'現在のGeneric BOM / 製品マッチング結果に実際に含まれる製品をサプライチェーン視点で再構成します。',need:'先に設計計算を実行してください。',center:'現在のBOM',used:'BOM項目',sources:'製品 / データシート',close:'閉じる',refresh:'更新',
      cats:{compute:'GPU / Accelerator',cpu:'Server CPU',network:'Network Fabric',optical:'Optical Connectivity',power:'Power / UPS / PDU',cooling:'Cooling / HVAC',rack:'Rack / Physical',storage:'Storage',monitoring:'Fiber Test / Monitoring',facility:'Facility / Security'}},
  zh:{button:'供应链',title:'BOM 供应链',sub:'将当前 Generic BOM / 产品匹配结果中实际包含的产品按供应链视角重新整理。',need:'请先执行设计计算。',center:'当前 BOM',used:'BOM 项目',sources:'产品 / 数据表',close:'关闭',refresh:'刷新',
      cats:{compute:'GPU / Accelerator',cpu:'Server CPU',network:'Network Fabric',optical:'Optical Connectivity',power:'Power / UPS / PDU',cooling:'Cooling / HVAC',rack:'Rack / Physical',storage:'Storage',monitoring:'Fiber Test / Monitoring',facility:'Facility / Security'}},
  de:{button:'Lieferkette',title:'BOM-Lieferkette',sub:'Ordnet die tatsächlich im aktuellen Generic BOM / Produkt-Matching enthaltenen Produkte als Lieferkettenansicht neu.',need:'Bitte zuerst die Designberechnung ausführen.',center:'Aktuelles BOM',used:'BOM-Positionen',sources:'Produkt / Datenblatt',close:'Schließen',refresh:'Aktualisieren',
      cats:{compute:'GPU / Accelerator',cpu:'Server CPU',network:'Network Fabric',optical:'Optical Connectivity',power:'Power / UPS / PDU',cooling:'Cooling / HVAC',rack:'Rack / Physical',storage:'Storage',monitoring:'Fiber Test / Monitoring',facility:'Facility / Security'}}
};

const VENDORS=[
  "NTest","M2 Optics","DSIT Solutions","FS (FiberStore)","Yokogawa","Moog","VIAVI","Anritsu","VeEX",
  'NVIDIA','AMD','Huawei','Qualcomm','Google Cloud','Google','Intel','Biren Technology','Ampere Computing','Juniper','Cisco','Arista','Broadcom','Credo','Lenovo','Supermicro','Dell','HPE','Hewlett Packard Enterprise',
  'Corning','SENKO','US Conec','ZTT','Sumitomo Electric','Sumitomo','LS Cable & System','LS Cable','YOFC','Hengtong','Lightera','Fujikura','Furukawa Electric','FITEL','Draka','Prysmian','Nexans','CommScope','Molex','Amphenol','Panduit','Belden','HYC','IH Optics','SHIJIA Photons','Hangzhou Zsine','Coherent','Lumentum','Marvell',
  'Schneider Electric','Schneider','APC','Vertiv','Eaton','Legrand','ABB','Rittal','Delta','Flex','LS ELECTRIC','MPOWERSYS','XEONICS','Green Power Technology','Gaon Cable','Taihan Cable & Solution','Taihan','Siemens','Generac','Caterpillar','CAT',
  'Pure Storage','NetApp','IBM','Micron','Samsung','SK hynix','Solidigm','Kioxia','KIOXIA','Western Digital','Phison','Marvell','Silicon Motion',
  'Mitsubishi Electric','MPS','Hitachi','Rolls-Royce','Atlas Copco','EnerSys','Cummins','Munters','STULZ','Carrier','Trane','Modine',
  'Fortinet','Palo Alto Networks','Palo Alto','Bosch','Securitas','Oracle','Fujitsu','Emerson','Asetek'
];
const CATEGORY_RULES=[
  ['monitoring',/Fiber Test|Monitoring|OTDR|FiberWatch|ONMSi|FTH-5000|928-OMS|RTU-4000|RTU-4100|NTest|M2 Optics|DSIT|Yokogawa|Moog|VIAVI|Anritsu|VeEX/i],
  ['cpu',/server\s*CPU|\bCPU\b|Xeon|EPYC|Grace CPU|AmpereOne|processor/i],
  ['compute',/\bGPU\b|\bNPU\b|\bTPU\b|Ascend|Atlas 900|Gaudi|AI200|AI250|Dragonfly|BR100|Ironwood|Trillium|Google Cloud|DGX|H100|H200|B200|B300|GB200|GB300|NVL72|MI300|MI350|MI355|Instinct|compute|server|accelerator|supermicro|lenovo|dell|hpe|hewlett/i],
  ['network',/switch|leaf|spine|core|fabric|NIC|DPU|ConnectX|BlueField|Spectrum|QFX|Nexus|Arista|Juniper|Cisco|Broadcom|Tomahawk|InfiniBand|Ethernet/i],
  ['optical',/optic|transceiver|fiber|fibre|cable|trunk|patch|breakout|harness|fan.?out|Fiber Shuffle|MPO|MTP|MMC|LC\b|OSFP|QSFP|AOC|DAC|AEC|DR4|FR4|SR8|Corning|Sumitomo|LS Cable|YOFC|Hengtong|Lightera|Fujikura|Furukawa|FITEL|fusion splicer|90S\+|90R|S179\+|S124M16|S185|Draka|Prysmian|Nexans|Credo|Coherent|Lumentum|Marvell|CommScope|Molex|Amphenol|Panduit|Belden|HYC|IH Optics|SHIJIA|Zsine/i],
  ['power',/UPS|PDU|power|power shelf|generator|transformer|busway|breaker|switchgear|Schneider|APC|Vertiv|Eaton|Legrand|ABB|Delta|Flex|LS ELECTRIC|MPOWERSYS|XEONICS|Green Power|Gaon Cable|Taihan|busduct|cable bus|MV cable|EHV cable|Generac|Caterpillar|\bCAT\b/i],
  ['cooling',/cooling|CDU|RDHx|chiller|HVAC|liquid|rear.?door|coolant|CRAC|CRAH|Vertiv|Schneider|Rittal|Munters|STULZ|Carrier|Trane|Modine|Emerson|Asetek/i],
  ['rack',/\brack\b|cabinet|enclosure|rail|cable manager|rack PDU|Rittal|Legrand/i],
  ['storage',/storage|NVMe|SSD|HDD|RAID|NAND|Pure Storage|NetApp|Solidigm|Kioxia|KIOXIA|Micron|Samsung|SK hynix|Phison|Marvell|Silicon Motion|Western Digital/i],
  ['facility',/security|fire|camera|access control|BMS|building automation|monitoring|sensor|Siemens|Fortinet|Palo Alto|Palo Alto Networks|Bosch|Securitas/i]
];
const RELATED_VENDORS={
  monitoring:["Fujikura", "NTest", "M2 Optics", "DSIT Solutions", "FS (FiberStore)", "Yokogawa", "Moog", "VIAVI", "Anritsu", "VeEX"],
  compute:['NVIDIA','AMD','Huawei','Intel','Qualcomm','Google Cloud','Biren Technology','Supermicro','Dell Technologies','HPE','Lenovo','Fujitsu','Oracle','IBM'],
  cpu:['Intel','AMD','NVIDIA','Ampere Computing'],
  network:['NVIDIA Networking','Broadcom','Arista Networks','Cisco','Juniper Networks','Marvell','HPE Aruba Networking'],
  optical:['Coherent','Lumentum','Corning','SENKO','US Conec','ZTT','Sumitomo Electric','Fujikura','Furukawa Electric','Lightera','Draka / Prysmian','Nexans','LS Cable & System','Hengtong','YOFC','HYC','IH Optics','SHIJIA Photons','Hangzhou Zsine','CommScope','Molex','Amphenol','Panduit','Belden'],
  network:[
    {vendor:'NVIDIA',product:'Spectrum-X / ConnectX',role:'AI Ethernet switch + NIC fabric',spec:'GPU-cluster Ethernet fabric · switch/NIC role separated',url:'https://www.nvidia.com/en-us/networking/'},
    {vendor:'Broadcom',product:'Tomahawk 5 / BCM78900',role:'Data-center Ethernet switch silicon',spec:'51.2 Tb/s · up to 64×800GbE / 128×400GbE',url:'https://www.broadcom.com/products/ethernet-connectivity/switching/strataxgs/bcm78900-series'},
    {vendor:'Arista',product:'7800R4 Series',role:'AI / cloud spine switching',spec:'Up to 460 Tb/s system throughput · up to 576×800GbE',url:'https://www.arista.com/ko/products/7800r4-series/specifications'}
  ],
  storage:[
    {vendor:'Samsung',product:'PM1743',role:'Enterprise PCIe 5.0 SSD',spec:'Up to 14,000 MB/s read · dual-port · enterprise / AI storage',url:'https://semiconductor.samsung.com/ssd/enterprise-ssd/pm1743/'},
    {vendor:'SK hynix',product:'PS1010 / PEB110',role:'AI / data-center eSSD',spec:'PCIe Gen5 · E3.S/U.2-U.3 and E1.S references',url:'https://news.skhynix.com/en/sk-hynix-develops-peb110-for-data-centers/'},
    {vendor:'KIOXIA',product:'CM9 Series',role:'Enterprise PCIe 5.0 SSD',spec:'PCIe 5.0 · NVMe 2.0 · E3.S · TLC · 1/3 DWPD families',url:'https://kr.kioxia.com/ko-kr/business/ssd/enterprise-ssd.html'},
    {vendor:'Micron',product:'9550 / 9650',role:'AI / data-center NVMe SSD',spec:'9550 PCIe Gen5 · 14 GB/s read · 9650 PCIe Gen6 portfolio reference',url:'https://www.micron.com/products/storage/ssd/data-center-ssd'},
    {vendor:'Phison',product:'Pascari D206V / Enterprise Controllers',role:'Enterprise SSD + NAND controller',spec:'D206V up to 245.76 TB · data-center scale',url:'https://www.phison.com/'},
    {vendor:'Marvell',product:'Bravera SC5',role:'Data-center SSD controller',spec:'PCIe 5.0 · NVMe 1.4b · 8/16 NAND channels · up to 14 GB/s',url:'https://www.marvell.com/products/ssd-controllers/mv-ss1331-1333.html'},
    {vendor:'Silicon Motion',product:'MonTitan SM8366 / SM8388',role:'Enterprise SSD controller',spec:'PCIe 5.0 · NVMe 2.0 · OCP DC NVMe · >14 GB/s · up to 128 TB reference',url:'https://www.siliconmotion.com/products/enterprise/detail'}
  ],
  power:['LS ELECTRIC','Schneider Electric','Vertiv','Eaton','Delta','Flex','MPOWERSYS','XEONICS','Green Power Technology','Gaon Cable','Taihan Cable & Solution','ABB','Siemens','Legrand','Mitsubishi Electric','Cummins','Caterpillar'],
  cooling:['Vertiv','Schneider Electric','Carrier','Trane','Munters','STULZ','Modine','Rittal','Asetek'],
  rack:['Rittal','Legrand','Vertiv','Eaton','Schneider Electric','HPE','Dell Technologies','Supermicro','Fujitsu'],
  storage:['Samsung','SK hynix','KIOXIA','Micron','Phison','Marvell','Silicon Motion','Solidigm','Pure Storage','NetApp','Dell Technologies','HPE','IBM','Western Digital'],
  facility:['Siemens','Schneider Electric','Fortinet','Palo Alto Networks','Bosch','Cisco','Securitas']
};

const VERIFIED_DC_PRODUCTS={
  monitoring:[{"vendor": "Fujikura", "product": "FlexScan FS200 OTDR", "role": "휴대형 OTDR · 설치/운영 장비 참고", "spec": "FTTH/PON · P2P 단일모드 광망", "url": "https://www.europe.fujikura.com/markets/telecoms/test-and-inspection/flexscan-fs200/"}, {"vendor": "NTest", "product": "FiberWatch Remote Fiber Test System", "role": "원격 광섬유 감시 시스템 (RFTS) · 설치/운영 장비 참고", "spec": "통신망 · DCI/캠퍼스 외부 광경로 · Dark/Live Fiber", "url": "https://www.ntestinc.com/fiberwatch-by-ntest"}, {"vendor": "M2 Optics", "product": "Dark and Lit Fiber Monitoring System", "role": "원격 광섬유 감시 시스템 · 설치/운영 장비 참고", "spec": "P2P · PON · Dark/Lit Fiber", "url": "https://www.m2optics.com/products/fiber-monitoring-systems"}, {"vendor": "FS (FiberStore)", "product": "FS FOTR-201 Handheld OTDR", "role": "휴대형 OTDR · 설치/운영 장비 참고", "spec": "광케이블 설치 · 현장 유지보수", "url": "https://www.fs.com/products/49652.html"}, {"vendor": "FS (FiberStore)", "product": "FS FMT Customized OTDR", "role": "FMT 플러그인 원격 OTDR · 설치/운영 장비 참고", "spec": "WDM/FMT 외부 광경로 · 플랫폼 호환 확인", "url": "https://www.fs.com/products/73281.html"}, {"vendor": "FS (FiberStore)", "product": "FS D7000 OTDR08", "role": "D7000 광경로 감시 모듈 · 설치/운영 장비 참고", "spec": "D7000 전송 플랫폼 광경로", "url": "https://www.fs.com/products/216495.html"}, {"vendor": "Yokogawa", "product": "Yokogawa AQ7277B Remote OTDR", "role": "원격 OTDR 모듈 · 설치/운영 장비 참고", "spec": "RFTS · 현용선 감시 · 장거리 광경로", "url": "https://tmi.yokogawa.com/kr/solutions/products/optical-measuring-instruments/optical-time-domain-reflectometer/aq7277-remote-optical-time-domain-reflectometer/"}, {"vendor": "Moog", "product": "Moog Focal 928-OMS Optical Monitoring System", "role": "산업/해양 광 텔레메트리 감시 · 설치/운영 장비 참고", "spec": "ROV · 해저 제어 · 중요 광 텔레메트리 (DCI 적용 별도 검토)", "url": "https://www.moog.com/products/multiplexers-media-converters/focal-multiplexer-product-line/condition-monitoring/model-928-oms.html"}, {"vendor": "VIAVI", "product": "VIAVI ONMSi Remote Fiber Test System", "role": "원격 광섬유 감시 시스템 (RFTS) · 설치/운영 장비 참고", "spec": "Core · Metro · FTTH/PON · DCI", "url": "https://www.viavisolutions.com/en-us/products/onmsi-remote-fiber-test-system-rfts"}, {"vendor": "VIAVI", "product": "VIAVI FTH-5000 Compact Fiber Test Head", "role": "랙 장착 원격 OTDR 테스트 헤드 · 설치/운영 장비 참고", "spec": "원격 광경로 감시 · DCI/캠퍼스 광망", "url": "https://www.viavisolutions.com/en-us/products/fth-5000"}, {"vendor": "VIAVI", "product": "VIAVI SmartOTDR Handheld Fiber Tester", "role": "휴대형 OTDR · 설치/운영 장비 참고", "spec": "Metro · Access · FTTH/PON · 단일모드 현장 검사", "url": "https://www.viavisolutions.com/en-us/products/smartotdr-handheld-fiber-tester"}, {"vendor": "Anritsu", "product": "Anritsu ACCESS Master MT9085 Series", "role": "휴대형 OTDR / 광손실 시험기 · 설치/운영 장비 참고", "spec": "Core/Metro · 이동통신 광경로 · Access 유지보수", "url": "https://www.anritsu.com/ko-kr/test-measurement/products/mt9085series"}, {"vendor": "VeEX", "product": "VeEX RTU-4000/4100 Remote Fiber Test System", "role": "원격 광섬유 테스트 시스템 (RFTS) · 설치/운영 장비 참고", "spec": "통신망 · 외부 광경로 원격 검사/감시", "url": "https://www.veexinc.com/products/remote-fiber-test-system-rfts-rtu-4000-4100"}],
  compute:[
    {vendor:'Huawei',product:'Atlas 900 A3 SuperPoD',role:'Ascend 910 NPU AI supernode',spec:'Up to 384 NPU · 128 GB/NPU · up to 3.2 TB/s memory bandwidth · liquid-cooled compute cabinets',url:'https://e.huawei.com/cn/products/computing/ascend/atlas-900-a3-superpod'},
    {vendor:'AMD',product:'Instinct MI350X / MI355X',role:'Physical AI/HPC accelerator',spec:'288 GB HBM3E · 8 TB/s · OEM/server platform validation required',url:'https://www.amd.com/en/products/accelerators/instinct/mi350.html'},
    {vendor:'Intel',product:'Gaudi 3',role:'AI training / inference accelerator',spec:'128 GB HBM2e · standard Ethernet scale-out',url:'https://www.intel.com/content/www/us/en/products/details/processors/ai-accelerators/gaudi.html'},
    {vendor:'Qualcomm',product:'Dragonfly AI200',role:'Rack-scale inference accelerator',spec:'768 GB/card · 56 cards/rack · 140 kW DLC rack',url:'https://www.qualcomm.com/data-center/products/qualcomm-dragonfly-ai200'}
  ],
  power:[
    {vendor:'Vertiv',product:'Liebert EXL S1',role:'Large / hyperscale data-center UPS',spec:'250–1200 kW family',url:'https://www.vertiv.com/en-us/products-catalog/critical-power/uninterruptible-power-supplies-ups/liebert-exl-s1/'},
    {vendor:'Eaton',product:'9395XR / 9395X',role:'AI / hyperscale modular UPS',spec:'Up to 1500 kW 9395XR · 9395X 900–1700 kW regional family',url:'https://www.eaton.com/us/en-us/catalog/backup-power-ups-surge-it-power-distribution/eaton-9395xr-ups.html'},
    {vendor:'Schneider Electric',product:'Galaxy VXL',role:'Large data-center modular UPS',spec:'500–1250 kW 400 V family · >97% efficiency reference',url:'https://www.se.com/kr/ko/work/products/product-launch/galaxy-vxl/'},
    {vendor:'Delta',product:'Modulon DPH',role:'Modular three-phase UPS',spec:'50–500 kVA · 220/380, 230/400, 240/415 V',url:'https://www.deltapowersolutions.com/ko-kr/mcis/50kw-500kw-three-phase-ups-dph-series-specifications.php'},
    {vendor:'Delta',product:'AI 120 kW / ORV3 18 kW Power Shelf',role:'AI 서버·랙 전원공급장치 · 시설 UPS와 구분',spec:'120 kW 총 정격 / 2N 보호 부하 60 kW · ORV3 1 OU 18 kW',url:'https://brandnews.deltaww.com/en/BrandCircleDetail/12483'},
    {vendor:'Flex',product:'CPRS / CRPS Server Power Supplies',role:'서버·스토리지·네트워크용 AC/DC PSU',spec:'이중화 서버 전원 제품군 · 정격/SKU별 공급 확인',url:'https://flex.com/downloads/power-cloud-server-solutions-ac-dc-power-supplies-for-data-centers'},
    {vendor:'Flex',product:'GB200 / GB300 Power Shelf',role:'Rack-level AI power shelf',spec:'1RU · 6 PSUs · up to 33 kW · Redfish monitoring',url:'https://flex.com/resources/power-shelves'},
    {vendor:'Gaon Cable',product:'AI Data Center Power Portfolio',role:'Korea · MV cable / Cable Bus / Busduct / EHV grid',spec:'US AI-data-center MV cable supply · CSA-certified cable bus · LSCUS busduct deployments · exact ratings RFQ',url:'https://www.gaoncable.com/en'},
    {vendor:'Taihan Cable & Solution',product:'Data Center Integrated Power Solution',role:'Korea · EHV / MV / LV cable + busduct',spec:'Project-tailored power infrastructure · PET/epoxy busduct portfolio · design-to-installation integration',url:'https://www.taihan.com/en/solutions/dataCenter'},
    {vendor:'LS ELECTRIC',product:'500 kVA UPS platform',role:'Korea · data-center UPS / MW parallel system reference',spec:'500 kVA unit · 5 units = 2.5 MW development system · 440 Vac · online · exact SKU RFQ',url:'https://nahpdev.ls-electric.com/markets/data-center'},
    {vendor:'MPOWERSYS',product:'MPS-3000 Series',role:'Korea · 3-phase True Online UPS',spec:'10–500 kVA · double conversion · input/output isolation transformer',url:'https://www.mpowersys.co.kr/bbs/page.php?hid=ups_mps03'},
    {vendor:'XEONICS',product:'XPS-NTI',role:'Korea · ALL-IGBT UPS',spec:'3-phase 10–500 kVA family · direct production',url:'https://www.xeonics.co.kr/50'},
    {vendor:'Green Power Technology',product:'GREEN UPS',role:'Korea · domestic high-efficiency UPS',spec:'1-phase 5–20 kVA · 3-phase 10–75 kVA · KS/KC stated',url:'https://www.greenups.co.kr/products-green.html'},
    {vendor:'Schneider Electric',product:'Galaxy VXL',role:'3-phase UPS · AI / large data center',spec:'500–1250 kW (400 V)',url:'https://www.se.com/kr/ko/product-range/209756733-galaxy-vxl/'},
    {vendor:'Eaton',product:'9395X UPS',role:'Hyperscale / colocation UPS',spec:'1.0–1.7 MVA · 97.5% online efficiency',url:'https://www.eaton.com/gb/en-gb/catalog/backup-power-ups-surge-it-power-distribution/eaton-9395x-ups.html'},
    {vendor:'ABB',product:'MegaFlex DPA',role:'High-density data-center UPS',spec:'250–1500 kW',url:'https://new.abb.com/ups/ups-and-power-conditioning/megaflex'},
    {vendor:'Vertiv',product:'Liebert APM2',role:'Modular mission-critical UPS',spec:'30–600 kVA · 400 V',url:'https://go.vertiv.com/LiebertAPM2'},
    {vendor:'Mitsubishi Electric',product:'9900D',role:'Hyperscale / colocation UPS',spec:'1200–2000 kVA · 480 V',url:'https://mitsubishicritical.com/uninterruptible-power-supplies/9900d/'},
    {vendor:'Siemens',product:'SIVACON S8',role:'LV power-distribution switchboard',spec:'IEC 61439-2 · data center / critical infrastructure',url:'https://www.siemens.com/en-us/products/sivacon/s8/'},
    {vendor:'Caterpillar',product:'C175-20',role:'Mission-critical / data-center generator',spec:'3150–4000 ekW · 60 Hz',url:'https://www.cat.com/en_US/products/new/power-systems/electric-power/diesel-generator-sets/1000028913.html'},
    {vendor:'Cummins',product:'QSK95 generator platform',role:'Data Center Continuous generator platform',spec:'C3500D5 example: 2500 kW DCC',url:'https://www.cummins.com/en-na/generators/products/qsk95'}
  ],
  optical:[
    {vendor:'Coherent',product:'800G DR8 / DR8+ Transceivers',role:'AI/cloud data-center optical transceivers',spec:'QSFP-DD/OSFP · 500 m to 2 km SMF product families',url:'https://www.coherent.com/networking/transceivers/datacom/FTCE4516E1PXM'},
    {vendor:'Lumentum',product:'800G / 1.6T Datacom Transceivers',role:'AI/cloud data-center optical transceivers',spec:'OSFP · 2×DR4 · up to 500 m SMF · 800G and 1.6T',url:'https://www.lumentum.com/en/products/data-center/datacom-transceivers'},
    {vendor:'Draka / Prysmian',product:'UCFUTURE M10',role:'Data-center high-density optical raceway cable',spec:'24F standard · 16F option · Ø5.3 mm · MPO/MTP · Cca-s1a-d1-a1',url:'https://www.prysmian.com/sites/www.prysmian.com/files/media/documents/M10_e_0.pdf'},
    {vendor:'Draka / Prysmian',product:'UCFIBRE D02b',role:'Data-center backbone / mini break-out cable',spec:'Up to 24F · ES9 tight buffer · FireRes LSHF-FR · duct/tray installation',url:'https://www.prysmian.com/en/en_multimedia_datacom_draka-ucfibre_indoor_tight_UCFIBRETM_I_Di_N_LSHF-FR_ES9_D02b.html'},
    {vendor:'Nexans',product:'LANmark-OF ENSPACE UHD',role:'Ultra-high-density data-center optical patching',spec:'Up to 144 LC or 72 MTP ports per 1U · 1U/2U/4U panels',url:'https://www.nexans.no/en/products/Data-Network-Solutions/Fibre-LAN-Systems/Fibre-patch-panels/LANmark-OF37292.html'},
    {vendor:'Nexans',product:'LANmark-OF Micro-Bundle Indoor',role:'Data-center indoor backbone cable',spec:'12/24/48/96F · LSZH · IEC 60332-1/-3 · splice/pigtail termination',url:'https://www.nexans.no/en/products/Data-Network-Solutions/Fibre-LAN-Systems/Fibre-cables/LANmark-OF37249.html'},
    {vendor:'Fujikura',product:'90S+ / 90R',role:'Fusion-splicing installation tools',spec:'90S+: core alignment single-fiber · 90R: mass/ribbon up to 16F',url:'https://www.fusionsplicer.fujikura.com/products/'},
    {vendor:'Furukawa Electric / FITEL',product:'S179+ / S124M16 / S185EDV',role:'Fusion-splicing installation & specialty-fiber tools',spec:'S179+ core alignment · S124M16 up to 16F ribbon · S185EDV supports PMF/MCF/LDF',url:'https://www.furukawaelectric.com/splicer/en/technical/'},
    {vendor:'HYC',product:'MPO Breakout Cable',role:'CIOE/OFC high-density breakout',spec:'8/12/16/24F · OM2/OM3/OM4/OM5 · low-loss IL ≤0.35 dB · -25~70 °C',url:'https://cn.hyc-system.com/Product/index_273/3308'},
    {vendor:'HYC',product:'MPO/MTP Branch Harness',role:'Data-center high-density harness',spec:'Multi-fiber main cable + sub-cables + breakout body + connectors · configurable',url:'https://cn.hyc-system.com/Product/index_273/1217'},
    {vendor:'Shenzhen IH Optics',product:'Data Center MPO/MTP Wiring System',role:'CIOE 2026 trunk / fan-out',spec:'MPO/MTP trunk + fan-out cable · panels · cassettes · LC · AOC/DAC · detailed specs RFQ',url:'https://exhibitors.cioe.cn/jtycn/cpen37839.html'},
    {vendor:'Hangzhou Zsine',product:'SM/MM MPO Optical Cable',role:'CIOE 2026 pre-terminated MPO',spec:'8–144 cores · 10G–400G · factory pre-terminated/tested · optional pulling grip',url:'https://exhibitors.cioe.cn/jtycn/cpen23025.html'},
    {vendor:'SHIJIA Photons',product:'MPO/MTP/MMC Assemblies & Fiber Shuffle',role:'OFC 2026 AI data-center interconnect',spec:'MPO/MTP/MMC assemblies · Fiber Shuffle · high-fiber-count cable up to 3,456F',url:'https://sjphotons.com/2026/03/23/ofc-2026-concluded-successfully-shijia-photons-global-partners-leveraging-light-to-explore-new-possibilities-in-the-ai-era/'}
  ],
  cooling:[
    {vendor:'Vertiv',product:'Liebert XDU450',role:'Coolant Distribution Unit',spec:'453 kW nominal · up to 975 kW max',url:'https://www.vertiv.com/en-us/products-catalog/thermal-management/high-density-solutions/liebert-xdu450-coolant-distribution-unit/'},
    {vendor:'Carrier',product:'AquaForce 30XF',role:'Mission-critical data-center air-cooled screw chiller',spec:'Integrated free-cooling platform',url:'https://www.carrier.com/us/en/commercial/chillers/'},
    {vendor:'Trane',product:'CenTraVac CDHH / CVHH',role:'Water-cooled data-center chiller',spec:'CDHH up to 21 MW · CVHH up to 9 MW',url:'https://www.trane.com/commercial/north-america/us/en/products-systems/chillers/data-center-chillers/centravac-data-center-chiller.html'}
  ],
  rack:[
    {vendor:'Vertiv',product:'VR Rack family',role:'Data-center rack / enclosure',spec:'VR3100 / VR3300 / VR3150 / VR3350 families',url:'https://www.vertiv.com/en-us/products-catalog/facilities-enclosures-and-racks/racks-and-containment/vertiv-rack/'}
  ],
  facility:[
    {vendor:'Fortinet',product:'FortiGate 3800G',role:'AI data-center firewall',spec:'200 Gbps threat protection · 400 GbE connectivity',url:'https://www.fortinet.com/solutions/data-center-firewall'}
  ]
};
window.__dcBomVerifiedSupplyCatalog=VERIFIED_DC_PRODUCTS;

function verifiedProductHtml(k){
  const arr=VERIFIED_DC_PRODUCTS[k]||[];
  if(!arr.length)return'';
  return '<div class="sc-verified"><div class="sc-related-title">'+(k==='monitoring'?'Fiber test / monitoring references <span>설치·운영 장비 · 적용 범위 확인</span>':'Verified DC products <span>공식 데이터센터 용도/사양 확인</span>')+'</div>'+
    arr.map(x=>'<div class="sc-vproduct"><div><b>'+esc(x.vendor)+'</b> · '+esc(x.product)+'</div><div class="sc-vrole">'+esc(x.role)+'</div><div class="sc-vspec">'+esc(x.spec)+'</div><a href="'+esc(x.url)+'" target="_blank" rel="noopener">Official product / datasheet</a></div>').join('')+
    '</div>';
}
function relatedHtml(k,used){
  const usedSet=new Set((used||[]).map(x=>(x.vendor||'').toLowerCase()).filter(Boolean));
  const arr=(RELATED_VENDORS[k]||[]).map(v=>typeof v==='string'?v:v.vendor).filter(v=>v&&!usedSet.has(v.toLowerCase()));
  const products=verifiedProductHtml(k);
  const companies=arr.length?'<div class="sc-related"><div class="sc-related-title">Related companies <span>참고 업체 · 현재 BOM 미선택</span></div><div class="sc-related-chips">'+arr.map(v=>'<span>'+esc(v)+'</span>').join('')+'</div></div>':'';
  return products+companies;
}

function token(s){s=String(s||'').trim().toLowerCase();if(s.includes('한국')||s.includes('korean')||s==='ko'||s==='kr')return'ko';if(s.includes('日本')||s.includes('japanese')||s==='ja'||s==='jp')return'ja';if(s.includes('中文')||s.includes('chinese')||s==='zh'||s==='cn')return'zh';if(s.includes('deutsch')||s.includes('german')||s==='de')return'de';if(s.includes('english')||s==='en')return'en';return null}
function lang(){if(window.__dcBomUiLang&&I18N[window.__dcBomUiLang])return window.__dcBomUiLang;const s=document.documentElement.getAttribute('data-dc-bom-ui-lang');if(s&&I18N[s])return s;return token(document.documentElement.lang)||'ko'}
const tr=()=>I18N[lang()]||I18N.ko;

function findCalc(){return Array.from(document.querySelectorAll('button,input[type="button"],input[type="submit"]')).find(e=>{const s=(e.value||textOf(e)).toLowerCase();return(s.includes('설계')&&s.includes('계산'))||s.includes('calculate')||s.includes('design calculation')||s.includes('計算')||s.includes('设计计算')||s.includes('berechnen')})||null}
function findTable(rx){
  for(const h of document.querySelectorAll('h1,h2,h3,h4,h5,summary,.phaseBand,.section-title,.card-title')){
    if(!rx.test(textOf(h)))continue;
    const scope=h.closest('.card,.panel,section,article,details')||h.parentElement;
    if(scope&&scope.querySelector('table'))return scope.querySelector('table');
    let n=h.nextElementSibling;
    for(let i=0;n&&i<8;i++,n=n.nextElementSibling){if(n.matches&&n.matches('table'))return n;if(n.querySelector&&n.querySelector('table'))return n.querySelector('table')}
  }
  return null;
}
function vendorFrom(text){const low=text.toLowerCase();return VENDORS.find(v=>low.includes(v.toLowerCase()))||''}
function categorize(text){for(const [key,rx] of CATEGORY_RULES)if(rx.test(text))return key;return'facility'}
function rowItems(table,source){
  if(!table)return[];
  const rows=Array.from(table.querySelectorAll('tr'));if(rows.length<2)return[];
  const headers=Array.from(rows[0].children).map(x=>textOf(x).toLowerCase());
  const idx=(rx)=>headers.findIndex(x=>rx.test(x));
  const vendorIdx=idx(/vendor|manufacturer|제조사|기업|maker|hersteller/);
  const productIdx=idx(/product|model|제품|모델|recommended|recommendation|추천/);
  const qtyIdx=idx(/qty|quantity|수량|menge/);
  return rows.slice(1).map((tr,ri)=>{
    const cells=Array.from(tr.children).filter(x=>x.tagName==='TD'||x.tagName==='TH');
    if(!cells.length)return null;
    const texts=cells.map(textOf),row=texts.join(' · ');
    if(!row||row==='-')return null;
    const links=Array.from(tr.querySelectorAll('a[href]')).map(a=>({label:textOf(a)||'Link',url:a.href})).filter(x=>/^https?:/i.test(x.url));
    const vendor=(vendorIdx>=0&&texts[vendorIdx])||vendorFrom(row);
    let product=(productIdx>=0&&texts[productIdx])||'';
    if(!product){product=texts.find((x,i)=>i!==vendorIdx&&i!==qtyIdx&&x&&x!=='-'&&!/^\d+(\.\d+)?$/.test(x))||row}
    const qty=qtyIdx>=0?texts[qtyIdx]:'';
    return{id:source+'-'+ri,source,vendor,product,qty,row,links,category:categorize(row)};
  }).filter(Boolean);
}
function collect(){
  const r=window.DCDesign;
  if(r){const primary=(r.products||[]).map((x,i)=>({id:'product-'+i,source:'receipt',vendor:x.vendor,product:x.product,qty:String(x.qty??''),row:[x.category,x.product,x.model].join(' · '),links:x.source&&/^https?:/i.test(x.source)?[{label:'Product / Datasheet',url:x.source}]:[],category:categorize([x.category,x.product,x.model].join(' · '))}));
    const secondary=(r.bom||[]).map((x,i)=>({id:'bom-'+i,source:'generic',vendor:'',product:x.item,qty:String(x.qty??''),row:[x.category,x.item].join(' · '),links:[],category:categorize([x.category,x.item].join(' · '))}));
    const seen=new Set(primary.map(x=>x.product.toLowerCase()));
    return [...primary,...secondary.filter(x=>!seen.has(x.product.toLowerCase()))];
  }
  const receipt=findTable(/제품.*매칭.*영수증|product.*match.*receipt|product.*receipt|製品.*マッチ|产品.*匹配|produkt.*matching/i);
  const generic=findTable(/generic\s*bom|일반\s*bom|범용\s*bom|汎用.*bom|通用.*bom/i);
  const primary=rowItems(receipt,'receipt'),secondary=rowItems(generic,'generic');
  const seen=new Set(),items=[];
  for(const x of [...primary,...secondary]){
    const k=(x.vendor+'|'+x.product+'|'+x.category).toLowerCase();
    if(seen.has(k))continue;seen.add(k);items.push(x);
  }
  return items;
}
function candidateProductsForRow(row){
  const r=String(row||'');
  if(/generator|genset|backup power|diesel|발전기/i.test(r))return VERIFIED_DC_PRODUCTS.power.filter(x=>/generator/i.test(x.role));
  if(/switchgear|switchboard|LV power|MV power|배전|MCC/i.test(r))return VERIFIED_DC_PRODUCTS.power.filter(x=>/switchboard/i.test(x.role));
  if(/UPS|uninterruptible|무정전/i.test(r))return VERIFIED_DC_PRODUCTS.power.filter(x=>/UPS/i.test(x.role)).slice(0,5);
  if(/CDU|coolant distribution|liquid cooling|cooling|HVAC|chiller|냉각/i.test(r))return VERIFIED_DC_PRODUCTS.cooling;
  if(/rack|cabinet|enclosure|랙|캐비닛/i.test(r))return VERIFIED_DC_PRODUCTS.rack;
  if(/firewall|security|보안/i.test(r))return VERIFIED_DC_PRODUCTS.facility;
  return[];
}
function enrichProductReceipt(){
  const table=findTable(/제품.*매칭.*영수증|product.*match.*receipt|product.*receipt|製品.*マッチ|产品.*匹配|produkt.*matching/i);
  if(!table)return false;
  const rows=Array.from(table.querySelectorAll('tr'));if(rows.length<2)return false;
  const heads=Array.from(rows[0].children).map(x=>textOf(x).toLowerCase());
  let ai=heads.findIndex(x=>/alternative|대안|대체|候補|替代|alternate/.test(x));
  if(ai<0)ai=heads.length-1;
  for(const tr of rows.slice(1)){
    if(tr.getAttribute('data-dc-supply-enriched')==='1')continue;
    const cells=Array.from(tr.children).filter(x=>x.tagName==='TD'||x.tagName==='TH');if(!cells[ai])continue;
    const cand=candidateProductsForRow(textOf(tr));if(!cand.length){tr.setAttribute('data-dc-supply-enriched','1');continue}
    const d=document.createElement('div');d.className='sc-receipt-candidates';
    d.innerHTML='<b>Verified DC candidates</b><br>'+cand.slice(0,5).map(x=>'<a href="'+esc(x.url)+'" target="_blank" rel="noopener">'+esc(x.vendor)+' · '+esc(x.product)+'</a>').join(' · ');
    cells[ai].appendChild(d);tr.setAttribute('data-dc-supply-enriched','1');
  }
  return true;
}

function esc(s){return String(s||'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]))}
function itemHtml(x){
  const vendor=x.vendor||'Vendor TBD';
  const qty=x.qty&&x.qty!=='-'?'<span class="sc-qty">× '+esc(x.qty)+'</span>':'';
  const src=x.source==='receipt'?'Product Match':'Generic BOM';
  const links=x.links.slice(0,2).map(a=>'<a href="'+esc(a.url)+'" target="_blank" rel="noopener">'+esc(a.label||'Product / Datasheet')+'</a>').join(' ');
  return'<div class="sc-item"><div class="sc-vendor">'+esc(vendor)+'<span class="sc-source">'+src+'</span></div><div class="sc-product">'+esc(x.product)+qty+'</div>'+(links?'<div class="sc-links">'+links+'</div>':'')+'</div>';
}
function modal(){
  let m=document.getElementById('supply-chain-modal');if(m)return m;
  m=document.createElement('div');m.id='supply-chain-modal';m.hidden=true;m.setAttribute('role','dialog');m.setAttribute('aria-modal','true');m.setAttribute('aria-label','Supply Chain');
  m.innerHTML='<div class="sc-shell"><div class="sc-top"><div><h2 class="sc-title"></h2><p class="sc-sub"></p></div><div class="sc-top-actions"><button type="button" class="sc-refresh"></button><button type="button" class="sc-close">×</button></div></div><div class="sc-body"></div><div class="sc-foot">v'+VERSION+' · Generated from the current on-screen BOM / product-match result</div></div>';
  document.body.appendChild(m);
  m.querySelector('.sc-close').onclick=close;
  m.querySelector('.sc-close').setAttribute('aria-label',tr().close);
  m.addEventListener('click',e=>{if(e.target===m)close()});
  m.querySelector('.sc-refresh').onclick=render;
  return m;
}
function render(){
  const m=modal(),t=tr(),items=collect();
  m.querySelector('.sc-title').textContent=t.title;
  m.querySelector('.sc-sub').textContent=t.sub;
  m.querySelector('.sc-refresh').textContent=t.refresh;
  
  const groups={};for(const k of Object.keys(t.cats))groups[k]=[];
  items.forEach(x=>(groups[x.category]||(groups[x.category]=[])).push(x));
  const vendorCount=new Set(items.map(x=>x.vendor).filter(Boolean)).size;
  const center='<section class="sc-center"><div class="sc-center-icon">DC</div><h3>'+t.center+'</h3><div class="sc-center-stat"><b>'+items.length+'</b> '+t.used+'</div><div class="sc-center-stat"><b>'+vendorCount+'</b> Vendors</div></section>';
  const order=['compute','cpu','network','optical','power','cooling','rack','storage','monitoring','facility'];
  const cards=order.filter(k=>selectedCategory==='all'||k===selectedCategory).map(k=>{
    const arr=groups[k]||[];
    const body=(arr.length?arr.map(itemHtml).join(''):'<div class="sc-none">현재 BOM 선택 업체 없음</div>')+relatedHtml(k,arr);
    return'<section class="sc-card sc-'+k+'"><div class="sc-cat-head"><span class="sc-dot"></span><h3>'+esc(t.cats[k])+'</h3><span class="sc-count">'+arr.length+' used</span></div>'+body+'</section>';
  }).join('');
  m.querySelector('.sc-body').innerHTML='<nav class="sc-filters" aria-label="부품별 업체"><button type="button" data-category="all" aria-pressed="'+(selectedCategory==='all')+'">전체</button>'+order.map(k=>'<button type="button" data-category="'+k+'" aria-pressed="'+(selectedCategory===k)+'">'+esc(t.cats[k])+'</button>').join('')+'</nav>'+(window.DCDesignStale?'<p class="sc-stale">입력값 변경됨 · BOM 수량은 이전 계산 기준입니다.</p>':'')+'<div class="sc-map'+(selectedCategory!=='all'?' sc-single':'')+'">'+(selectedCategory==='all'?center:'')+cards+'</div>';
  m.querySelectorAll('[data-category]').forEach(b=>b.onclick=()=>{selectedCategory=b.dataset.category;render();m.querySelector('[data-category="'+selectedCategory+'"]')?.focus();});
}
function close(){const m=document.getElementById('supply-chain-modal');if(m)m.hidden=true;returnFocus?.focus();}
function open(){returnFocus=document.activeElement;render();modal().hidden=false;modal().querySelector('.sc-close').focus();}
function mount(){
  if(document.getElementById('supply-chain-btn'))return true;
  const calc=document.getElementById('run')||findCalc();if(!calc||!calc.parentNode)return false;
  const host=document.createElement('span');host.id='supply-chain-inline-host';
  const b=document.createElement('button');b.type='button';b.id='supply-chain-btn';b.textContent=tr().button;b.title='BOM 설계에 실제 사용된 업체/제품 보기';b.onclick=open;host.appendChild(b);
  calc.insertAdjacentElement('afterend',host);
  return true;
}
function localize(){const b=document.getElementById('supply-chain-btn');if(b)b.textContent=tr().button;const m=document.getElementById('supply-chain-modal');if(m&&!m.hidden)render()}
function start(){
  let tourSupplyOpened=false;
  const openTourSupply=()=>{if(!tourSupplyOpened&&location.hash==='#supply-chain'&&document.getElementById('supply-chain-btn')){tourSupplyOpened=true;open();}};
  mount();openTourSupply();
  document.addEventListener('keydown',e=>{const m=document.getElementById('supply-chain-modal');if(!m||m.hidden)return;if(e.key==='Escape')close();if(e.key==='Tab'){const a=[...m.querySelectorAll('button,a[href]')],first=a[0],last=a[a.length-1];if(e.shiftKey&&document.activeElement===first){e.preventDefault();last.focus();}else if(!e.shiftKey&&document.activeElement===last){e.preventDefault();first.focus();}}});
  document.addEventListener('dc:design',()=>{localize();enrichProductReceipt();});
  let attempts=0;const boot=setInterval(()=>{attempts++;const mounted=mount();openTourSupply();const enriched=enrichProductReceipt();if((mounted&&enriched)||attempts>24)clearInterval(boot)},500);
  document.addEventListener('click',e=>{const q=e.target&&e.target.closest&&e.target.closest('[data-lang],[data-language],button,a,[role="button"]');if(q)setTimeout(()=>{localize();enrichProductReceipt()},50)},true);
  document.addEventListener('change',()=>setTimeout(()=>{localize();enrichProductReceipt()},30),true);
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start,{once:true});else start();
})();
