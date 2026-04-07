// ============================================================
//  ESIC Settlement Form Filler  —  esic_filler.js
//  Hosted on GitHub Pages. Loaded dynamically by the loader
//  bookmarklet. Update this file to push changes to all users.
//
//  Field IDs confirmed from live portal: 06-Apr-2026
//  Sections: Travel, Boarding & Lodging, Local Conveyance, Misc
// ============================================================

(function() {
  // TEMPLATE_VERSION — bump this string whenever the Excel template changes.
  // Users uploading old templates will be prompted to download the latest.
  const TEMPLATE_VERSION = 'v3';
  const TEMPLATE_URL = 'https://viveki1989.github.io/esic-settlement-filler/ESIC_Settlement_Template.xlsx';

  if (!location.href.includes('gateway.esic.gov.in') && !location.href.includes('esic.gov.in')) {
    alert('Please click this bookmarklet from the ESIC Gateway page (gateway.esic.gov.in)');
    return;
  }
  const OLD = document.getElementById('_esicBMPanel');
  if (OLD) { OLD.remove(); return; }

  function loadXLSX(cb) {
    if (window.XLSX) { cb(); return; }
    const s = document.createElement('script');
    s.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    s.onload = cb; document.head.appendChild(s);
  }

  const sleep = ms => new Promise(r => setTimeout(r, ms));

  function normDate(raw) {
    if (!raw) return '';
    raw = String(raw).trim();
    if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(raw)) return raw;
    if (!isNaN(raw) && Number(raw) > 40000) {
      const d = new Date(Math.round((Number(raw) - 25569) * 86400 * 1000));
      const p = n => String(n).padStart(2,'0');
      return `${p(d.getUTCDate())}/${p(d.getUTCMonth()+1)}/${d.getUTCFullYear()}`;
    }
    return raw;
  }

  // Combine separate date and time into "DD/MM/YYYY HH:MM" for ESIC form
  function combineDateTime(datePart, timePart) {
    const d = normDate(datePart);
    if (!d) return '';
    const t = String(timePart||'').trim();
    // Normalise time: accept H:MM, HH:MM, HHMM, or Excel decimal fraction
    let timeStr = '00:00';
    if (t) {
      if (/^\d{1,2}:\d{2}$/.test(t)) {
        // HH:MM or H:MM — pad hour
        const [h,m] = t.split(':');
        timeStr = String(h).padStart(2,'0') + ':' + String(m).padStart(2,'0');
      } else if (/^\d{3,4}$/.test(t)) {
        // HHMM format
        timeStr = t.slice(0,-2).padStart(2,'0') + ':' + t.slice(-2);
      } else if (!isNaN(t) && Number(t) > 0 && Number(t) < 1) {
        // Excel time as decimal fraction (e.g. 0.291666 = 07:00)
        const totalMins = Math.round(Number(t) * 24 * 60);
        const hh = Math.floor(totalMins / 60);
        const mm = totalMins % 60;
        timeStr = String(hh).padStart(2,'0') + ':' + String(mm).padStart(2,'0');
      } else {
        timeStr = t.slice(0,5); // fallback — take first 5 chars
      }
    }
    return `${d} ${timeStr}`;
  }

  function parseWorkbook(xlsxWb) {
    let travelReqNo = '';
    const ws0 = xlsxWb.Sheets['HOW_TO_USE'];
    if (ws0) {
      // Read C3 directly: r=2 (0-based), c=2 (col C)
      const cell = ws0[XLSX.utils.encode_cell({r:2, c:2})];
      travelReqNo = cell ? String(cell.v || cell.w || '').trim() : '';
      // Fallback: scan row 3 for any non-label value
      if (!travelReqNo) {
        for (let c=0; c<=5; c++) {
          const cl = ws0[XLSX.utils.encode_cell({r:2, c})];
          if (cl) {
            const v = String(cl.v||cl.w||'').trim();
            if (v && !v.includes('→') && !v.startsWith('⚠') && !v.startsWith('Enter') && !v.startsWith('Travel')) {
              travelReqNo = v; break;
            }
          }
        }
      }
    }
    function section(name, c1, c2, r0) {
      const ws = xlsxWb.Sheets[name]; if (!ws) return [];
      const data = XLSX.utils.sheet_to_json(ws, {header:1,defval:'',raw:true});
      const rows = [];
      for (let i = r0; i < Math.min(data.length, r0+20); i++) {
        const r = (data[i]||[]).slice(c1, c2+1);
        if (r.every(v => v===''||v===0||v===null)) continue;
        rows.push(r);
      }
      return rows;
    }
    // Travel: 13 cols (A-M), data from Excel row 4 = array index 3
    // D=startDate E=startTime F=endDate G=endTime (split for user, combined for form)
    const tv = section('Travel_Settlement',1,12,3).map(r=>({
      from:      String(r[0]||'').trim(),
      to:        String(r[1]||'').trim(),
      startDate: combineDateTime(r[2], r[3]),  // D+E → DD/MM/YYYY HH:MM
      endDate:   combineDateTime(r[4], r[5]),  // F+G → DD/MM/YYYY HH:MM
      relaxation:String(r[6]||'').trim().toLowerCase()==='yes',
      mode:      String(r[7]||'').trim(),
      classCoach:String(r[8]||'').trim(),
      bookedBy:  String(r[9]||'').trim(),
      ticketNo:  String(r[10]||'').trim(),
      amount:    r[11]||0
    })).filter(r=>r.from||r.to);
    // B&L: 11 cols (A-K), data from Excel row 4 = array index 3
    // E=fromDate F=fromTime G=toDate H=toTime (split for user, combined for form)
    const bl = section('Boarding_Lodging',1,10,3).map(r=>({
      blType:  String(r[0]||'').trim(),           // lodgingBoardingType
      accType: String(r[1]||'').trim(),            // accomodationType
      bookedBy:String(r[2]||'').trim(),            // accomodationBookedBy
      fromDate:combineDateTime(r[3], r[4]),        // E+F → DD/MM/YYYY HH:MM
      toDate:  combineDateTime(r[5], r[6]),        // G+H → DD/MM/YYYY HH:MM
      days:    r[7]||'',                           // absenceDays
      billNo:  String(r[8]||'').trim(),            // billNo
      actual:  r[9]||0                             // claimAmount
    })).filter(r=>r.blType||r.fromDate);
    // LC: confirmed IDs lcFromLocation{n}, lcToLocation{n}, lcFromDate{n}, lcToDate{n},
    //   lcModeofTravel{n}, noOfKms{n}, ratePerKM{n}, lcActualAmt{n}
    const lc = section('Local_Conveyance',1,8,2).map(r=>({
      from:    String(r[0]||'').trim(),   // lcFromLocation
      to:      String(r[1]||'').trim(),   // lcToLocation
      fromDate:normDate(r[2]),            // lcFromDate
      toDate:  normDate(r[3]),            // lcToDate
      mode:    String(r[4]||'').trim(),   // lcModeofTravel
      km:      r[5]||0,                   // noOfKms
      rate:    r[6]||0,                   // ratePerKM
      actual:  r[7]||0                    // lcActualAmt
    })).filter(r=>r.from||r.to);
    // Misc: confirmed IDs expenseType{n}, miscellAmount{n} — NO bill no field on portal
    const ms = section('Miscellaneous',1,2,2).map(r=>({
      desc:  String(r[0]||'').trim(),     // expenseType
      actual:r[1]||0                      // miscellAmount
    })).filter(r=>r.desc);
    // Version detection: v3 template has 13 cols in Travel_Settlement
    // (split date/time). Old templates have 11. Detect by checking col E header.
    let detectedVersion = 'v3'; // assume current
    const wsTv = xlsxWb.Sheets['Travel_Settlement'];
    if (wsTv) {
      // Row 3 (index 2), col E (index 4) should contain "Time" in v3
      const hdrE = wsTv[XLSX.utils.encode_cell({r:2, c:4})];
      const hdrEval = hdrE ? String(hdrE.v||'').toLowerCase() : '';
      if (!hdrEval.includes('time')) detectedVersion = 'old';
    }

    return { travelReqNo, travelRows:tv, blRows:bl, lcRows:lc, miscRows:ms, detectedVersion };
  }

  function setField(w, id, val) {
    const el = w.document.getElementById(id); if (!el) return false;
    try {
      const proto = el.tagName==='SELECT' ? w.HTMLSelectElement.prototype : w.HTMLInputElement.prototype;
      const desc = Object.getOwnPropertyDescriptor(proto,'value');
      if (desc&&desc.set) desc.set.call(el,String(val)); else el.value=String(val);
    } catch(e){el.value=String(val);}
    ['input','change','blur'].forEach(ev=>el.dispatchEvent(new w.Event(ev,{bubbles:true})));
    return true;
  }
  function setSelect(w, id, text) {
    const el = w.document.getElementById(id); if (!el||el.tagName!=='SELECT') return false;
    const t = String(text).trim().toLowerCase();
    const opt = [...el.options].find(o=>o.text.trim().toLowerCase()===t)
             || [...el.options].find(o=>o.value.trim().toLowerCase()===t);
    if (!opt) return false;
    return setField(w, id, opt.value);
  }
  function setCheck(w, id, checked) {
    const el = w.document.getElementById(id); if (!el) return false;
    if (el.checked!==!!checked){el.checked=!!checked;el.dispatchEvent(new w.Event('change',{bubbles:true}));}
    return true;
  }
  function waitFor(w, sel, ms=15000) {
    return new Promise((res,rej)=>{
      const t0=Date.now();
      const f=()=>{
        try{const e=w.document.querySelector(sel);if(e){res(e);return;}}catch(e){}
        if(Date.now()-t0>ms){rej(new Error('Timeout: '+sel));return;}
        setTimeout(f,400);
      };f();
    });
  }

  async function fillForm(w, data, log) {
    const {travelRows:tv, blRows:bl, lcRows:lc, miscRows:ms} = data;

    // ── Travel ──────────────────────────────────────────────────────────────
    log('Filling Travel Settlement…','info');
    for (let i=0; i<Math.max(0,tv.length-2); i++) {
      try{w.insertTASettlementRow();}catch(e){}
      await sleep(400);
    }
    for (let i=0; i<tv.length; i++) {
      const r=tv[i]; log(`  Row ${i+1}: ${r.from} → ${r.to}`);
      setField(w,`tsFromLocation${i}`,r.from);
      setField(w,`tsToLocation${i}`,r.to);
      setField(w,`tsFromDate${i}`,r.startDate);
      setField(w,`tsToDate${i}`,r.endDate);
      setCheck(w,`tsRelaxation${i}`,r.relaxation);
      setSelect(w,`tsModeofTravel${i}`,r.mode);
      await sleep(200);
      setSelect(w,`tsClassCoach${i}`,r.classCoach);
      setSelect(w,`tsBookedBy${i}`,r.bookedBy);
      setField(w,`tsTicketNo${i}`,r.ticketNo);
      setField(w,`ticketAmount${i}`,r.amount);
    }
    log(`  ✓ ${tv.length} row(s)`,'ok');

        // ── B&L — all IDs hardcoded from live portal ────────────────────────────
    // insertAccomodationSettlementRow() confirmed
    // IDs: lodgingBoardingType{n} → triggers dependent accomodationType{n}
    //      accomodationBookedBy{n}, accomodationFromDate{n}, accomodationToDate{n}
    //      absenceDays{n}, billNo{n}, claimAmount{n}
    if (bl.length) {
      log('Filling Boarding & Lodging…','info');
      for (let i=0; i<bl.length; i++) {
        try{w.insertAccomodationSettlementRow();}catch(e){}
        await sleep(500);
        const r=bl[i]; const n=i;
        log(`  B&L ${i+1}: ${r.blType} – ${r.accType}`);
        setSelect(w, `lodgingBoardingType${n}`, r.blType);
        await sleep(700); // let dependent dropdown populate
        setSelect(w, `accomodationType${n}`,    r.accType);
        setSelect(w, `accomodationBookedBy${n}`,r.bookedBy);
        setField(w,  `accomodationFromDate${n}`,r.fromDate);
        setField(w,  `accomodationToDate${n}`,  r.toDate);
        setField(w,  `absenceDays${n}`,          r.days);
        setField(w,  `billNo${n}`,               r.billNo);
        setField(w,  `claimAmount${n}`,          r.actual);
      }
      log(`  ✓ ${bl.length} B&L row(s)`,'ok');
    } else log('B&L: no rows','skip');

        // ── LC — all IDs hardcoded from live portal ─────────────────────────────
    // insertLocalConveyanceRow() confirmed
    // IDs: lcFromLocation{n}, lcToLocation{n}, lcFromDate{n}, lcToDate{n},
    //      lcModeofTravel{n}, noOfKms{n}, ratePerKM{n}, lcActualAmt{n}
    if (lc.length) {
      log('Filling Local Conveyance…','info');
      for (let i=0; i<lc.length; i++) {
        try{w.insertLocalConveyanceRow();}catch(e){}
        await sleep(500);
        const r=lc[i]; const n=i;
        log(`  LC ${i+1}: ${r.from} → ${r.to}`);
        setField(w,  `lcFromLocation${n}`, r.from);
        setField(w,  `lcToLocation${n}`,   r.to);
        setField(w,  `lcFromDate${n}`,      r.fromDate);
        setField(w,  `lcToDate${n}`,        r.toDate);
        setSelect(w, `lcModeofTravel${n}`,  r.mode);
        setField(w,  `noOfKms${n}`,         r.km);
        setField(w,  `ratePerKM${n}`,       r.rate);
        setField(w,  `lcActualAmt${n}`,     r.actual);
      }
      log(`  ✓ ${lc.length} LC row(s)`,'ok');
    } else log('LC: no rows','skip');

        // ── Misc — all IDs hardcoded from live portal ───────────────────────────
    // insertMiscellRow() confirmed
    // IDs: expenseType{n}, miscellAmount{n}  — NO bill no field on portal
    if (ms.length) {
      log('Filling Miscellaneous…','info');
      for (let i=0; i<ms.length; i++) {
        try{w.insertMiscellRow();}catch(e){}
        await sleep(500);
        const r=ms[i]; const n=i;
        log(`  Misc ${i+1}: ${r.desc}`);
        setField(w, `expenseType${n}`,   r.desc);
        setField(w, `miscellAmount${n}`, r.actual);
      }
      log(`  ✓ ${ms.length} misc row(s)`,'ok');
    } else log('Misc: no rows','skip');
  }

  async function runFlow(data, log, setProg) {
    const {travelReqNo:trNo} = data;
    if (!trNo) { log('❌ Travel Request No. missing in Excel (HOW_TO_USE sheet, row 3, col C)','err'); return false; }
    log(`Travel Request No.: ${trNo}`,'info');

    // Step 1: find HRMS 2.0 onclick link
    log('Step 1: Opening HRMS 2.0…','info'); setProg(8);
    let hrmsOnclick = null;
    document.querySelectorAll('a[onclick]').forEach(a=>{
      const oc=a.getAttribute('onclick')||'';
      if (oc.includes('ESICHRMSV2ClientNew') && !oc.includes('Support') && !oc.includes('FAQ'))
        if (a.textContent.trim()==='HRMS 2.0') hrmsOnclick=oc;
    });
    if (!hrmsOnclick) { log('❌ HRMS 2.0 link not found on page','err'); return false; }
    const urlM = hrmsOnclick.match(/window\.open\('([^']+)'/);
    if (!urlM) { log('❌ Could not parse HRMS 2.0 URL','err'); return false; }

    const hw = window.open(urlM[1],'_esicHRMS','width=1280,height=820,left=60,top=40,scrollbars=yes,resizable=yes');
    if (!hw) { log('❌ Popup blocked — allow popups for gateway.esic.gov.in and try again','err'); return false; }
    log('Step 1: HRMS 2.0 window opened, waiting for SSO…','info'); setProg(15);

    // Wait for HRMS domain
    const ready = await new Promise(res=>{
      const t0=Date.now();
      const f=()=>{
        try{ if(hw.location.href.includes('ESICHRMSV2')){res(true);return;} }catch(e){}
        if(Date.now()-t0>50000){res(false);return;}
        setTimeout(f,700);
      };f();
    });
    if (!ready) { log('❌ HRMS 2.0 did not load in 50s — check SSO / VPN','err'); return false; }
    await sleep(1800);
    log('Step 2: HRMS loaded ✓','ok'); setProg(30);

    // Step 3: navigate to TR list
    log('Step 3: Going to Travel Request List…','info');
    const base = hw.location.origin+'/ESICHRMSV2';
    hw.location.href = `${base}/hrmsTransactions/travelRequest_travelRequestList.action?time=${Date.now()}`;
    await sleep(2800);
    try { await waitFor(hw,'#myRequest',12000); } catch(e) { log('❌ TR List page did not load','err'); return false; }
    log('Step 3: TR List loaded ✓','ok'); setProg(45);

    // Step 4: filter — only My Request = Yes (no Request Type filter)
    log('Step 4: Setting My Request = Yes…','info');
    setSelect(hw,'myRequest','Yes'); await sleep(300);
    try{hw.showTRList('');}catch(e){
      hw.document.querySelectorAll('input[type=button]').forEach(b=>{ if(b.value.trim()==='Get Details') b.click(); });
    }
    await sleep(2800);
    log('Step 4: My Request = Yes applied ✓','ok'); setProg(58);

    // Step 5: find TR No link
    log(`Step 5: Looking for ${trNo}…`,'info');
    let trLink = null;
    const findLink = () => {
      hw.document.querySelectorAll('a').forEach(a=>{
        if (a.textContent.trim()===trNo || a.href.includes(trNo)) trLink=a;
      });
      if (!trLink) {
        hw.document.querySelectorAll('td[onclick], tr[onclick]').forEach(el=>{
          if ((el.textContent||'').includes(trNo)) trLink=el;
        });
      }
    };
    findLink(); if (!trLink) { await sleep(2000); findLink(); }
    if (!trLink) { log(`❌ "${trNo}" not found in list — check if it exists as Settlement type under My Request=Yes`,'err'); return false; }
    log(`Step 5: Found ${trNo} ✓`,'ok'); setProg(68);
    trLink.click();

    // Step 6: wait for settlement page
    log('Step 6: Waiting for settlement page…','info');
    const settled = await new Promise(res=>{
      const t0=Date.now();
      const f=()=>{
        try{if(hw.location.href.includes('settlementView.action')){res(true);return;}}catch(e){}
        if(Date.now()-t0>25000){res(false);return;}
        setTimeout(f,500);
      };f();
    });
    if (!settled) { log('❌ Settlement page did not load','err'); return false; }
    await sleep(2500);
    try { await waitFor(hw,'#tsFromLocation0',10000); } catch(e) { log('❌ Form fields not found on settlement page','err'); return false; }
    log('Step 6: Settlement page loaded ✓','ok'); setProg(78);

    // Step 7: fill
    log('Step 7: Filling form…','info');
    await fillForm(hw, data, log);
    setProg(100);
    log('━━━━━━━━━━━━━━━━━━━━━━━━━━━','info');
    log('✅ Done! Review and click Submit.','ok');
    try{hw.focus();}catch(e){}
    return true;
  }

  // ── Panel styles ──────────────────────────────────────────────────────────
  const sty = document.createElement('style');
  sty.textContent = `
    #_esicBMPanel *{box-sizing:border-box}
    @keyframes _bmi{from{transform:translateX(100%)}to{transform:translateX(0)}}
    #_esicBMPanel{position:fixed;top:0;right:0;width:375px;height:100vh;background:#fff;
      border-left:1px solid #E2E8F0;box-shadow:-8px 0 36px rgba(0,0,0,.15);
      z-index:2147483647;font-family:'Segoe UI',Arial,sans-serif;
      display:flex;flex-direction:column;animation:_bmi .22s ease}
    ._ph{background:#1B2A4A;color:#fff;padding:15px 18px;display:flex;
      justify-content:space-between;align-items:center;flex-shrink:0}
    ._ph h2{font-size:14.5px;font-weight:600;margin:0}
    ._ph span{font-size:11px;color:#93C5FD;display:block;margin-top:2px}
    ._pc{background:rgba(255,255,255,.15);border:none;color:#fff;width:27px;height:27px;
      border-radius:50%;cursor:pointer;font-size:17px;display:flex;align-items:center;justify-content:center}
    ._pb{padding:16px;overflow-y:auto;flex:1;display:flex;flex-direction:column;gap:0}
    ._tip{background:#FFF7ED;border-left:3px solid #F59E0B;border-radius:0 6px 6px 0;
      padding:8px 12px;font-size:11.5px;color:#92400E;line-height:1.5;margin-bottom:12px}
    ._uz{border:2px dashed #CBD5E1;border-radius:10px;padding:22px 14px;text-align:center;
      cursor:pointer;background:#F8FAFC;margin-bottom:11px;transition:border-color .2s,background .2s}
    ._uz:hover,._uz.drag{border-color:#2563EB;background:#EFF6FF}
    ._uz p{font-size:13px;color:#475569;margin:7px 0 0}
    ._uz .s{font-size:11px;color:#94A3B8;margin-top:3px}
    #_bfi{display:none}
    ._ch{background:#EFF6FF;border:1px solid #BFDBFE;border-radius:7px;padding:9px 12px;
      font-size:12px;color:#1E40AF;margin-bottom:11px;display:none;align-items:center;gap:8px}
    ._pw{height:5px;background:#E2E8F0;border-radius:3px;margin:8px 0 11px;display:none;overflow:hidden}
    ._pf{height:100%;width:0%;background:#2563EB;border-radius:3px;transition:width .4s}
    ._sb{width:100%;background:#1B2A4A;color:#fff;border:none;border-radius:9px;padding:12px;
      font-size:14px;font-weight:600;cursor:pointer;display:flex;align-items:center;
      justify-content:center;gap:8px;transition:background .15s}
    ._sb:hover{background:#2563EB}._sb:disabled{background:#94A3B8;cursor:not-allowed}
    ._lg{display:none !important}
    ._ll{margin:2px 0}
    ._ok{color:#4ADE80}._er{color:#F87171}._sk{color:#64748B}._in{color:#93C5FD}._df{color:#CBD5E1}
    ._dn{background:#F0FDF4;border:1px solid #BBF7D0;border-radius:8px;padding:11px 13px;
      font-size:12px;line-height:1.6;display:none;margin-top:10px}
    ._fb{margin-top:auto;padding:14px 0 4px;font-size:11px;color:#64748B;text-align:center;border-top:1px solid #E2E8F0}
    ._fb a{color:#2563EB;text-decoration:none}
    ._fb a:hover{text-decoration:underline}
    ._note{background:#F8FAFC;border:1px solid #E2E8F0;border-left:3px solid #2563EB;border-radius:0 8px 8px 0;padding:10px 13px;font-size:11px;color:#475569;line-height:1.6;margin-bottom:10px}
    ._note strong{color:#1B2A4A;font-size:11.5px}
    ._vmm{background:#FFF7ED;border:1px solid #FED7AA;border-radius:8px;padding:10px 13px;font-size:12px;color:#92400E;display:none;margin-bottom:10px}
    ._vmm a{color:#B45309;font-weight:600;text-decoration:underline;cursor:pointer}
  `;
  document.head.appendChild(sty);

  const panel = document.createElement('div');
  panel.id = '_esicBMPanel';
  panel.innerHTML = `
    <div class="_ph">
      <div><h2>⚡ ESIC Settlement Filler</h2><span>Upload Excel → navigate → fill form automatically</span></div>
      <button class="_pc" id="_bcp">×</button>
    </div>
    <div class="_pb">
      <div class="_note"><strong>Note &nbsp;·&nbsp;</strong> This tool processes travel data only within your Excel file on your device. It does not transmit or store any data externally, does not collect credentials, and is intended solely to assist in filling travel settlement details in the ESIC Gateway.</div>
      <div class="_vmm" id="_bvmm">⚠ You appear to be using an older Excel template. <a id="_bdl" href="#" target="_blank">Download the latest template</a> and re-upload before proceeding.</div>
      <div class="_tip">Stay on the Gateway page. Upload filled Excel — bookmarklet handles the rest.</div>
      <div class="_uz" id="_buz">
        <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#94A3B8" stroke-width="1.5">
          <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/>
          <polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/>
        </svg>
        <p>Click to choose Excel file or drag &amp; drop</p>
        <p class="s">ESIC_Settlement_Template.xlsx</p>
        <input type="file" id="_bfi" accept=".xlsx,.xls">
      </div>
      <div class="_ch" id="_bch"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#16A34A" stroke-width="2.5"><polyline points="20 6 9 17 4 12"/></svg><span id="_bfn">—</span></div>
      <div class="_pw" id="_bpw" style="display:none"><div class="_pf" id="_bpf"></div></div>
      <button class="_sb" id="_bsb" disabled>
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="5 3 19 12 5 21 5 3"/></svg>
        Start — open HRMS &amp; fill form
      </button>
      <div class="_lg" id="_blg"></div>
      <div class="_dn" id="_bdn"></div>
      <div class="_fb">For any feedback/suggestions/queries, drop a mail at <a href="mailto:vivek.i@esic.gov.in">vivek.i@esic.gov.in</a></div>
    </div>`;
  document.body.appendChild(panel);

  document.getElementById('_bcp').onclick = () => panel.remove();

  const zone=document.getElementById('_buz'), fi=document.getElementById('_bfi'),
    ch=document.getElementById('_bch'), fn=document.getElementById('_bfn'),
    sb=document.getElementById('_bsb'), lg=document.getElementById('_blg'),
    dn=document.getElementById('_bdn'), pw=document.getElementById('_bpw'),
    pf=document.getElementById('_bpf');

  let pdata=null;
  zone.onclick=()=>fi.click();
  zone.ondragover=e=>{e.preventDefault();zone.classList.add('drag');};
  zone.ondragleave=()=>zone.classList.remove('drag');
  zone.ondrop=e=>{e.preventDefault();zone.classList.remove('drag');go(e.dataTransfer.files[0]);};
  fi.onchange=()=>go(fi.files[0]);

  function go(file){
    if(!file)return;
    ch.style.display='flex'; lg.style.display='none'; lg.innerHTML=''; dn.style.display='none';
    loadXLSX(()=>{
      const rd=new FileReader();
      rd.onload=e=>{
        try{
          const wb=XLSX.read(new Uint8Array(e.target.result),{type:'array'});
          pdata=parseWorkbook(wb);
          const{travelReqNo:tr,travelRows:tv,blRows:bl,lcRows:lc,miscRows:ms,detectedVersion:dv}=pdata;
          // Show version mismatch warning if old template detected
          const vmm=document.getElementById('_bvmm');
          const bdl=document.getElementById('_bdl');
          if(dv==='old'){
            vmm.style.display='block';
            bdl.href=TEMPLATE_URL;
          } else {
            vmm.style.display='none';
          }
          fn.innerHTML=tr
            ?`<strong style="color:#15803D">${file.name}</strong> &nbsp; TR: <strong>${tr}</strong> · ${tv.length} travel · ${bl.length} B&L · ${lc.length} LC · ${ms.length} misc`
            :`<span style="color:#DC2626">⚠ TR No. missing — fill HOW_TO_USE sheet row 3 col C</span>`;
          sb.disabled=!tr;
          if(!tr)sb.innerHTML='⚠ Enter Travel Request No. in Excel';
          else sb.innerHTML='<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="5 3 19 12 5 21 5 3"/></svg> Start — open HRMS &amp; fill form';
        }catch(err){fn.innerHTML=`<span style="color:#DC2626">❌ ${err.message}</span>`;}
      };
      rd.readAsArrayBuffer(file);
    });
  }

  function addLog(msg,type='df'){
    lg.style.display='block';
    const d=document.createElement('div'); d.className=`_ll _${type}`;
    const t=new Date().toLocaleTimeString('en-IN',{hour12:false,hour:'2-digit',minute:'2-digit',second:'2-digit'});
    d.textContent=`[${t}] ${msg}`; lg.appendChild(d); lg.scrollTop=lg.scrollHeight;
  }
  function setProg(p){pf.style.width=p+'%';pw.style.display='block';}

  sb.onclick=async function(){
    if(!pdata)return;
    sb.disabled=true; sb.textContent='⏳ Running…';
    lg.innerHTML=''; dn.style.display='block'; dn.innerHTML='<span style="color:#64748B">⏳ Opening HRMS and navigating… please wait.</span>'; setProg(3);
    let ok=false;
    try{ ok=await runFlow(pdata,addLog,setProg); }
    catch(err){ addLog('Fatal: '+err.message,'er'); }
    sb.disabled=false;
    sb.innerHTML='<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="5 3 19 12 5 21 5 3"/></svg> Run again';
    dn.style.display='block';
    if(ok){
      dn.innerHTML=`<strong style="color:#15803D;font-size:13px">✅ Form filled successfully!</strong><br><span style="color:#374151;font-size:12px">Please switch to the HRMS window, review all sections carefully, then click <strong>Submit</strong>.</span>`;
      // Inject a visible banner on the settlement page itself
      try{
        const hw=window.open('','_esicHRMS');
        if(hw && hw.document && hw.document.body){
          const existing=hw.document.getElementById('_esicFilledBanner');
          if(existing) existing.remove();
          const banner=hw.document.createElement('div');
          banner.id='_esicFilledBanner';
          banner.style.cssText='position:fixed;top:0;left:0;width:100%;z-index:99999;background:#166534;color:#fff;font-family:Arial,sans-serif;font-size:14px;font-weight:600;text-align:center;padding:12px 20px;box-shadow:0 2px 12px rgba(0,0,0,.25);display:flex;align-items:center;justify-content:center;gap:12px';
          banner.innerHTML='<span style="font-size:18px">✅</span> Form filled automatically. Please check all sections carefully before clicking Submit. <button onclick="this.parentElement.remove()" style="margin-left:16px;background:rgba(255,255,255,.2);border:1px solid rgba(255,255,255,.4);color:#fff;padding:4px 12px;border-radius:5px;cursor:pointer;font-size:12px">Dismiss</button>';
          hw.document.body.insertBefore(banner,hw.document.body.firstChild);
          hw.focus();
        }
      }catch(e){}
    } else {
      dn.innerHTML=`<strong style="color:#DC2626;font-size:13px">⚠ Completed with some issues.</strong><br><span style="color:#374151;font-size:12px">The form may be partially filled. Please review all sections before submitting.</span>`;
    }
  };
})();
