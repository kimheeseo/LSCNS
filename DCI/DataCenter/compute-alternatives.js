(() => {
  'use strict';

  const ROWS = [
    {
      category: 'AI Accelerator Alternative',
      item: 'Huawei Atlas 900 A3 SuperPoD · Ascend 910 NPU',
      qty: 'Reference',
      unit: 'platform',
      basis: 'Physical AI infrastructure candidate · Huawei-specific SuperPoD, power, liquid cooling and networking design required'
    },
    {
      category: 'AI Accelerator Alternative',
      item: 'Intel Gaudi 3',
      qty: 'Reference',
      unit: 'candidate',
      basis: 'Physical accelerator candidate · 128 GB HBM2e · Ethernet scale-out · OEM/platform validation required'
    },
    {
      category: 'AI Accelerator Alternative',
      item: 'Qualcomm Dragonfly AI200',
      qty: 'Reference',
      unit: 'rack/candidate',
      basis: 'Inference-focused accelerator · 768 GB/card · 56-card / 140 kW rack reference · platform validation required'
    },
    {
      category: 'AI Accelerator Alternative',
      item: 'Biren BR100 family',
      qty: 'Reference',
      unit: 'candidate',
      basis: 'Regional accelerator supplier reference · availability, compliance and platform specifications require RFQ'
    },
    {
      category: 'AI Accelerator Alternative',
      item: 'AMD Instinct MI350X / MI355X',
      qty: 'Reference',
      unit: 'candidate',
      basis: 'Physical accelerator BOM candidate · OEM/server platform power, cooling and fabric validation required before quantity substitution'
    },
    {
      category: 'Server CPU Alternative',
      item: 'AMD EPYC 9005 Series',
      qty: 'Reference',
      unit: 'candidate',
      basis: 'Physical server CPU candidate · OEM platform/socket/memory configuration validation required'
    },
    {
      category: 'Cloud AI Accelerator',
      item: 'Google Cloud TPU7x (Ironwood) / TPU v6e (Trillium)',
      qty: 'Reference',
      unit: 'capacity',
      basis: 'Cloud capacity / reservation reference only · excluded from physical procurement quantity'
    }
  ];

  function addReferenceRows() {
    const body = document.getElementById('bomBody');
    if (!body || !body.children.length) return;

    // Colocation-only scenarios do not use an accelerator/server BOM.
    if (window.DCDesign?.input?.scenario === 'colo') {
      body.querySelectorAll('[data-compute-alt-row]').forEach(el => el.remove());
      return;
    }

    ROWS.forEach((row, idx) => {
      const key = 'compute-alt-' + idx;
      if (body.querySelector('[data-compute-alt-row="' + key + '"]')) return;
      const tr = document.createElement('tr');
      tr.setAttribute('data-compute-alt-row', key);
      [row.category, row.item, row.qty, row.unit, row.basis].forEach(value => {
        const td = document.createElement('td');
        td.textContent = value;
        tr.appendChild(td);
      });
      body.appendChild(tr);
    });

    const panel = body.closest('.panel');
    if (panel && !document.getElementById('compute-alt-note')) {
      const note = document.createElement('p');
      note.id = 'compute-alt-note';
      note.className = 'muted';
      note.textContent = 'AMD/Huawei/Intel/Qualcomm/Biren/TPU 항목은 비교·조달 검토용 Reference입니다. 검증된 vendor-specific 서버·rack 프로파일이 추가되기 전까지 현재 NVIDIA 기반 rack/power/network 계산값은 임의로 변경하지 않습니다.';
      panel.appendChild(note);
    }

    try {
      if (window.currentDesignSnapshot) {
        window.currentDesignSnapshot.results = window.currentDesignSnapshot.results || {};
        window.currentDesignSnapshot.results.computeAlternatives = ROWS.map(x => ({...x, referenceOnly:true}));
      }
    } catch (_) {}
  }

  function start() {
    const body = document.getElementById('bomBody');
    if (!body) {
      setTimeout(start, 250);
      return;
    }

    const observer = new MutationObserver(() => queueMicrotask(addReferenceRows));
    observer.observe(body, { childList: true });
    document.addEventListener('dc:design', addReferenceRows);
    document.addEventListener('click', () => setTimeout(addReferenceRows, 120), true);
    addReferenceRows();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', start, { once: true });
  } else {
    start();
  }
})();
