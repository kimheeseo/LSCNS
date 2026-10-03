/* Dates are curated from primary company publications, not scraped in-browser. */
(() => {
 const checked='2026-10-03';
 const companies=[
 ['NVIDIA','https://www.nvidia.com/gtc/events/','확정 일정 수록'],
 ['AMD','https://ir.amd.com/news-events/ir-calendar','공식 IR + 기술발표 일정'],
 ['LS전선','https://news.lscns.com/','공식 뉴스 · 향후 확정일 미확인'],
 ['Corning','https://investor.corning.com/news-and-events/events-and-presentations/default.aspx','동적 IR 페이지 · 향후 확정일 미확인'],
 ['SENKO','https://www.senko.com/events/','공식 목록에서 향후 확정일 미확인'],
 ['US Conec','https://www.usconec.com/about-us/events','현재 게시된 2026 행사 중 향후 확정일 미확인'],
 ['ZTT','https://www.zttgroup.com/events/','공식 행사 페이지 조회 제한 · 날짜 미확인'],
 ['YOFC','https://en.yofc.com/list/210.html','공식 전시 뉴스 · 향후 확정일 미확인'],
 ['Hengtong','https://www.hengtonggroup.com/en/home/news/index/categoryId/3.html','공식 전시 목록 · 향후 확정일 미확인'],
 ['Sumitomo Electric','https://sumitomoelectric.com/ir/calendar','공식 IR + 전시 일정'],
 ['Fujikura','https://www.europe.fujikura.com/events/2026-events/','공식 기술·전시 일정'],
 ['Furukawa Electric','https://www.furukawaelectric.com/product/exhibition/','공식 전시 + IR 자료 · 향후 DC 관련 확정일 미확인']
 ].map(([company,source,status])=>({company,source,status,checked}));
 const entries=[
 ['LS전선','Data Centre World Asia · 신제품 공개','2026-09-29','2026-09-30','전시·기술','Singapore','Asia/Singapore','https://news.lscns.com/ls%EC%A0%84%EC%84%A0-ai-%EB%8D%B0%EC%9D%B4%ED%84%B0%EC%84%BC%ED%84%B0%EC%9A%A9-%EC%8B%A0%EC%A0%9C%ED%92%88-%EA%B3%B5%EA%B0%9C%EC%84%9C%EB%B2%84%EB%9E%99-%EB%82%B4%EB%B6%80%EB%A1%9C-%EC%82%AC/'],
 ['NVIDIA','GTC Berlin','2026-10-20','2026-10-22','Summit','Berlin','Europe/Berlin','https://www.nvidia.com/gtc/events/'],
 ['NVIDIA','GTC Washington, D.C.','2026-11-30','2026-12-03','Summit','Washington, D.C.','America/New_York','https://www.nvidia.com/gtc/dc/'],
 ['NVIDIA','GTC 2027','2027-03-15','2027-03-18','Summit','공식 공지 확인','현지','https://www.nvidia.com/gtc/events/'],
 ['AMD','Lisa Su keynote · OCP Global Summit','2026-10-12','2026-10-12','기술발표','OCP Global Summit','현지','https://newsroom.amd.com/news/media-alert-ceo-lisa-su-keynote-2026-ocp/'],
 ['AMD','PyTorch Conference','2026-10-20','2026-10-21','기술행사','San Jose','America/Los_Angeles','https://www.amd.com/en/corporate/events.html'],
 ['AMD','Advancing AI 2026','2026-07-23','2026-07-23','기술발표','공식 발표','현지','https://ir.amd.com/news-events/ir-calendar'],
 ['SENKO','ECOC Exhibition','2026-09-21','2026-09-23','전시·기술','Málaga','Europe/Madrid','https://www.senko.com/events/'],
 ['SENKO','CIOE','2026-09-09','2026-09-11','전시·기술','Shenzhen','Asia/Shanghai','https://www.senko.com/events/'],
 ['US Conec','SCTE TechExpo · booth L1803','2026-09-29','2026-10-01','전시·기술','Atlanta','America/New_York','https://www.usconec.com/about-us/events'],
 ['US Conec','ECOC · stand 1216','2026-09-21','2026-09-23','전시·기술','Málaga','Europe/Madrid','https://www.usconec.com/about-us/events'],
 ['YOFC','OFC 2026','2026-03-15','2026-03-19','전시·기술','Los Angeles','America/Los_Angeles','https://en.yofc.com/view/3542.html'],
 ['Sumitomo Electric','Quarterly results announcement','2026-10-30','2026-10-30','IR','Japan','Asia/Tokyo','https://sumitomoelectric.com/ir/calendar'],
 ['Sumitomo Electric','VISION 2026','2026-10-06','2026-10-08','전시·기술','Stuttgart','Europe/Berlin','https://sumitomoelectric.com/jp/company/exhibitions'],
 ['Fujikura','Capacity Europe','2026-10-13','2026-10-15','Summit','London','Europe/London','https://www.europe.fujikura.com/events/2026-events/'],
 ['Fujikura','Electronica','2026-11-10','2026-11-13','전시·기술','Munich','Europe/Berlin','https://www.europe.fujikura.com/events/2026-events/'],
 ['Furukawa Electric','Photonix 2026','2026-09-30','2026-10-02','전시·기술','Japan','Asia/Tokyo','https://www.furukawaelectric.com/product/exhibition/'],
 ['Furukawa Electric','Q1 results briefing','2026-08-06','2026-08-06','IR','Japan','Asia/Tokyo','https://www.furukawaelectric.com/en/ir/library/index.html']
 ].map(([company,title,start,end,type,location,timezone,source])=>({company,title,start,end,type,location,timezone,source,checked,time:'미공개 · 날짜 기준'}));
 function filter({company='',type='',past=false,from='',to='',today=new Date().toISOString().slice(0,10)}={}){return entries.filter(e=>(!company||e.company===company)&&(!type||e.type===type)&&(past||e.end>=today)&&(!from||e.end>=from)&&(!to||e.start<=to)).sort((a,b)=>a.start.localeCompare(b.start)||a.company.localeCompare(b.company));}
 window.DCOfficialEvents={checked,companies,entries,filter};
})();
