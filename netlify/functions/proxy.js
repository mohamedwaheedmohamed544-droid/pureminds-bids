/**
 * PUREMINDS BD — Netlify Serverless Proxy
 * ----------------------------------------
 * Fetches SharePoint Excel files server-side (no CORS issues),
 * parses them, extracts hyperlinks, and returns structured JSON.
 *
 * Endpoints:
 *   GET /.netlify/functions/proxy?type=tenders
 *   GET /.netlify/functions/proxy?type=krassat
 */

const XLSX = require('xlsx');

// ── Configure your SharePoint sharing URLs here ──────────────────────────
const SHARE_URLS = {
  tenders : process.env.URL_TENDERS  || 'PASTE_TENDERS_SHAREPOINT_URL_HERE',
  krassat : process.env.URL_KRASSAT  || 'PASTE_KRASSAT_SHAREPOINT_URL_HERE',
};
// ─────────────────────────────────────────────────────────────────────────

// Cache: avoid hitting SharePoint on every request (TTL = 5 minutes)
const cache = {};
const CACHE_TTL = 5 * 60 * 1000;

// ── Helpers ───────────────────────────────────────────────────────────────
const g   = v => (v == null ? '' : String(v).trim());
const fix = url => {
  if (!url) return '';
  if (url.startsWith('http')) return url;
  return 'https://puremindss-my.sharepoint.com/' + url.replace(/^(\.\.\/)+/, '');
};

function getCellLink(ws, r, c) {
  const addr = XLSX.utils.encode_cell({ r, c });
  return fix(ws[addr]?.l?.Target || '');
}

// ── Fetch the Excel binary from SharePoint ────────────────────────────────
async function fetchExcel(shareUrl) {
  const b64 = Buffer.from(shareUrl)
    .toString('base64')
    .replace(/=/g, '')
    .replace(/\+/g, '-')
    .replace(/\//g, '_');

  // Method 1: SharePoint REST API → get @microsoft.graph.downloadUrl (CDN link, no auth needed)
  try {
    const tenant = shareUrl.match(/https:\/\/([^\/]+)/)?.[1];
    if (tenant) {
      const metaUrl = `https://${tenant}/_api/v2.0/shares/u!${b64}/root?$select=@microsoft.graph.downloadUrl`;
      const metaRes = await fetch(metaUrl, {
        redirect: 'follow',
        headers: { 'Accept': 'application/json', 'User-Agent': 'Mozilla/5.0' },
      });
      if (metaRes.ok) {
        const meta = await metaRes.json();
        const dlUrl = meta['@microsoft.graph.downloadUrl'];
        if (dlUrl) {
          const fileRes = await fetch(dlUrl, { redirect: 'follow' });
          if (fileRes.ok) return Buffer.from(await fileRes.arrayBuffer());
        }
      }
    }
  } catch (_) {}

  // Method 2: Microsoft Graph API shares endpoint
  try {
    const metaUrl = `https://graph.microsoft.com/v1.0/shares/u!${b64}/driveItem?$select=@microsoft.graph.downloadUrl`;
    const metaRes = await fetch(metaUrl, {
      redirect: 'follow',
      headers: { 'Accept': 'application/json', 'User-Agent': 'Mozilla/5.0' },
    });
    if (metaRes.ok) {
      const meta = await metaRes.json();
      const dlUrl = meta['@microsoft.graph.downloadUrl'];
      if (dlUrl) {
        const fileRes = await fetch(dlUrl, { redirect: 'follow' });
        if (fileRes.ok) return Buffer.from(await fileRes.arrayBuffer());
      }
    }
  } catch (_) {}

  // Method 3: OneDrive consumer API
  try {
    const r = await fetch(`https://api.onedrive.com/v1.0/shares/u!${b64}/root/content`, {
      redirect: 'follow', headers: { 'User-Agent': 'Mozilla/5.0' },
    });
    if (r.ok) return Buffer.from(await r.arrayBuffer());
  } catch (_) {}

  // Method 4: Direct &download=1 (last resort)
  const dlUrl = shareUrl.includes('?') ? shareUrl + '&download=1' : shareUrl + '?download=1';
  const r4 = await fetch(dlUrl, { redirect: 'follow', headers: { 'User-Agent': 'Mozilla/5.0' } });
  if (!r4.ok) throw new Error(`SharePoint HTTP ${r4.status}`);
  const finalBuf = await r4.arrayBuffer();
  // Reject if response is HTML (means login redirect)
  const ct = r4.headers.get('content-type') || '';
  if (ct.includes('text/html')) throw new Error('SharePoint returned a login page — the file may not be publicly accessible');
  return Buffer.from(finalBuf);
}

// ── Parse منافسات 2026 ────────────────────────────────────────────────────
function parseTenders(buf) {
  const wb = XLSX.read(buf, { type: 'buffer', cellDates: true, cellStyles: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { defval: '', header: 1 });

  const hdr = (raw[0] || []).map(h => g(h));
  const idx = name => hdr.findIndex(h => h.includes(name));

  const iE   = idx('اسم الجهة'),      iT  = idx('نوع المنافسة');
  const iN   = idx('اسم المناقصة'),   iR  = idx('الرقم المرجعي');
  const iD   = idx('تاريخ تقديم'),    iTm = idx('اخر وقت');
  const iV   = idx('قيمة المناقصة'), iP  = idx('المنصة');
  const iO   = idx('رقم العرض المالي'), iDec = idx('تقييم');
  const iRep = idx('التقرير التحليلي'), iAn = idx('متابعة');
  const iNt  = idx('مهمة'),           iPrc = idx('تحليل الأسعار');
  const iPrc2= idx('تحليل الاسعار تندرز'), iCtc = idx('التواصل');
  const iFin = idx('الإعتماد النهائي'), iTF = idx('متابعة مهام الفريق');
  // Dept statuses
  const iIntr = idx('مقدمة العرض الفني'), iCnt = idx('قسم المحتوى');
  const iDig  = idx('قسم التسويق'),       iPR  = idx('قسم العلاقات العامة');
  const iTech = idx('التقنية'),           iAth = idx('تصاميم أثير');
  const iDes  = idx('قسم التصميم'),       iEv  = idx('إدارة الفعاليات');
  const iDE   = idx('تعديل تصاميم'),     iCr  = idx('تصميم الافكار');
  const iRev  = idx('مراجعة العرض');
  // Link columns
  const iLA = idx('المرفقات'),   iLS = idx('نطاق العمل');
  const iLR = idx('التقرير التحليلي'), iLT = idx('متابعة مهام الفريق');
  const iLC = idx('قسم المحتوى'), iLD = idx('قسم التسويق');
  const iLP = idx('قسم العلاقات'), iLTe = idx('التقنية');
  const iLDes = idx('قسم التصميم'), iLEv = idx('إدارة الفعاليات');

  const data = [];
  raw.slice(1).forEach((row, ri) => {
    const entity = g(row[iE]);
    if (!entity) return;

    let deadline = '';
    const dv = row[iD];
    if (dv instanceof Date) deadline = dv.toISOString().slice(0, 10);
    else if (dv) deadline = String(dv).slice(0, 10);

    let time = '';
    const tv = row[iTm];
    if (tv instanceof Date) time = tv.toTimeString().slice(0, 5);
    else if (typeof tv === 'string') time = tv.slice(0, 5);

    const excelRow = ri + 1; // 0-indexed raw → xlsx row index for getCellLink

    data.push({
      entity,
      type           : g(row[iT]),
      name           : g(row[iN]),
      ref            : g(row[iR]),
      deadline,
      time,
      value          : parseFloat(String(row[iV] || 0).replace(/[^\d.]/g, '')) || 0,
      platform       : g(row[iP]),
      offer_num      : g(row[iO]),
      decision       : g(row[iDec]),
      report         : g(row[iRep]),
      analyst        : g(row[iAn]),
      notes          : g(row[iNt]),
      price_analysis : g(row[iPrc]) || g(row[iPrc2]),
      contact        : g(row[iCtc]),
      final_approval : row[iFin] === true || row[iFin] === 1,
      team_followup  : g(row[iTF]),
      // Department statuses
      intro          : g(row[iIntr]),
      content_dept   : g(row[iCnt]),
      digital_dept   : g(row[iDig]),
      pr_dept        : g(row[iPR]),
      tech_dept      : g(row[iTech]),
      designs_atheer : g(row[iAth]),
      design_dept    : g(row[iDes]),
      events_mgmt    : g(row[iEv]),
      design_edit    : g(row[iDE]),
      creative       : g(row[iCr]),
      review         : g(row[iRev]),
      // Hyperlinks
      link_attachments : getCellLink(ws, excelRow, iLA),
      link_scope       : getCellLink(ws, excelRow, iLS),
      link_price       : getCellLink(ws, excelRow, iPrc) || getCellLink(ws, excelRow, iPrc2),
      link_report      : getCellLink(ws, excelRow, iLR),
      link_team        : getCellLink(ws, excelRow, iLT),
      link_content     : getCellLink(ws, excelRow, iLC),
      link_digital     : getCellLink(ws, excelRow, iLD),
      link_pr          : getCellLink(ws, excelRow, iLP),
      link_tech        : getCellLink(ws, excelRow, iLTe),
      link_design      : getCellLink(ws, excelRow, iLDes),
      link_events      : getCellLink(ws, excelRow, iLEv),
    });
  });

  return data;
}

// ── Parse متابعة الكراسات ─────────────────────────────────────────────────
function parseKrassat(buf) {
  const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });
  const g2 = v => (v == null ? '' : String(v).trim());
  const toBool = v => v === true || v === 1 || String(v).toUpperCase() === 'TRUE';
  const toN = v => { try { return parseFloat(String(v || 0).replace(/[^\d.]/g, '')) || 0; } catch { return 0; } };

  const findHdr = (rows, text) => {
    const i = rows.findIndex(r => r.some(c => String(c || '').includes(text)));
    return i < 0 ? 0 : i;
  };
  const mkIdx = hdrs => k => hdrs.findIndex(h => String(h || '').includes(k));

  const result = { reports: [], platforms: [], alerts: [], inquiries: [], daily: {}, financial: [] };

  // تقرير فتح العروض 2026
  const rptSheet = wb.SheetNames.find(s => s.includes('تقرير فتح العروض'));
  if (rptSheet) {
    const ws = wb.Sheets[rptSheet];
    const raw = XLSX.utils.sheet_to_json(ws, { defval: '', header: 1 });
    const hdrRow = findHdr(raw, 'الجهة');
    result.reports = raw.slice(hdrRow + 1).filter(r => g2(r[2]) && g2(r[2]) !== 'الجهة').map(r => ({
      entity        : g2(r[2]),
      name          : g2(r[3]),
      ref           : g2(r[4]),
      period        : g2(r[1]),
      platform      : g2(r[6]),
      analyst       : g2(r[7]),
      offer_num     : g2(r[8]),
      offer_value   : toN(r[10]),
      tech_status   : g2(r[11]),
      fin_status    : g2(r[12]),
      report_status : g2(r[13]),
      notes         : g2(r[14]),
      analysis      : g2(r[15]),
      has_boq       : toBool(r[17]),
      has_tech      : toBool(r[18]),
      has_report    : toBool(r[19]),
      review_status : g2(r[20]),
    })).filter(r => r.entity);
  }

  // العروض الماليه
  const finSheet = wb.SheetNames.find(s => s.includes('العروض الماليه'));
  if (finSheet) {
    const ws  = wb.Sheets[finSheet];
    const raw = XLSX.utils.sheet_to_json(ws, { defval: '', header: 1 });
    raw.slice(1).forEach((r, ri) => {
      const entity = g2(r[0]);
      if (!entity) return;
      let date = '';
      const dv = r[3];
      if (dv instanceof Date) date = dv.toISOString().slice(0, 10);
      else if (dv && String(dv).length > 3) date = String(dv).slice(0, 10);
      result.financial.push({
        entity,
        name       : g2(r[2]),
        date,
        file_name  : g2(r[6]),
        status     : g2(r[7]),
        notes      : g2(r[8]),
        link_file  : getCellLink(ws, ri + 1, 6),
      });
    });
  }

  // متابعة المنصات المعتمدة
  const pltSheet = wb.SheetNames.find(s => s.includes('المنصات المعتمدة'));
  if (pltSheet) {
    const ws  = wb.Sheets[pltSheet];
    const raw = XLSX.utils.sheet_to_json(ws, { defval: '', header: 1 });
    const hdr = findHdr(raw, 'اسم المنصة');
    result.platforms = raw.slice(hdr + 1).filter(r => g2(r[1])).map(r => ({
      num: r[0], name: g2(r[1]), notes: g2(r[3]), responsible: g2(r[4]),
    })).filter(r => !r.name.includes('اسم المنصة'));
  }

  // الاشعارات
  const alrtSheet = wb.SheetNames.find(s => s.trim().includes('الاشعارات'));
  if (alrtSheet) {
    const ws  = wb.Sheets[alrtSheet];
    const raw = XLSX.utils.sheet_to_json(ws, { defval: '', header: 1 });
    const hdr = findHdr(raw, 'اسم الجهة');
    const idx = mkIdx(raw[hdr] || []);
    result.alerts = raw.slice(hdr + 1).filter(r => g2(r[idx('اسم الجهة')] || r[0])).map(r => ({
      entity      : g2(r[idx('اسم الجهة')]    || r[0]),
      name        : g2(r[idx('اسم المناقصة')] || r[1]),
      ref         : g2(r[idx('الرقم المرجعي')] || r[2]),
      deadline    : g2(r[idx('تاريخ')]         || r[3]),
      platform    : g2(r[idx('المنصة')]         || r[4]),
      alert       : g2(r[idx('الإشعارات')]    || r[5]),
      responsible : g2(r[idx('المسؤول')]       || r[6]),
    })).filter(r => r.entity && !r.entity.includes('اسم الجهة'));
  }

  // الاستفسارات
  const inqSheet = wb.SheetNames.find(s => s.trim().includes('الاستفسارات'));
  if (inqSheet) {
    const ws  = wb.Sheets[inqSheet];
    const raw = XLSX.utils.sheet_to_json(ws, { defval: '', header: 1 });
    const hdr = findHdr(raw, 'اسم الجهة');
    const idx = mkIdx(raw[hdr] || []);
    result.inquiries = raw.slice(hdr + 1).filter(r => g2(r[idx('اسم الجهة')] || r[0])).map(r => ({
      entity      : g2(r[idx('اسم الجهة')]      || r[0]),
      name        : g2(r[idx('اسم المناقصة')]   || r[1]),
      ref         : g2(r[idx('الرقم المرجعي')]  || r[2]),
      deadline    : g2(r[idx('تاريخ')]           || r[3]),
      platform    : g2(r[idx('المنصة')]           || r[4]),
      inquiry     : g2(r[idx('الاستفسار')]       || r[5]),
      responsible : g2(r[idx('المسؤول')]         || r[6]),
    })).filter(r => r.entity && !r.entity.includes('اسم الجهة'));
  }

  // متابعة المهام اليومي
  const dailySheet = wb.SheetNames.find(s => s.includes('المهام اليوم'));
  if (dailySheet) {
    const ws  = wb.Sheets[dailySheet];
    const raw = XLSX.utils.sheet_to_json(ws, { defval: '', header: 1 });
    const week = g2(raw[1]?.[0] || '');
    const DAYS = ['الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];
    const days = raw.slice(4).filter(r => DAYS.some(d => String(r[0] || '').includes(d)));
    result.daily = {
      week,
      rows: days.map(r => ({
        day: g2(r[0]), responsible: g2(r[1]),
        tasks: [r[2], r[5], r[8], r[11], r[14]].map(v => toBool(v)),
      })),
    };
  }

  return result;
}

// ── Lambda handler ────────────────────────────────────────────────────────
exports.handler = async event => {
  const type = event.queryStringParameters?.type || 'tenders';

  if (!SHARE_URLS[type] || SHARE_URLS[type].startsWith('PASTE_')) {
    return {
      statusCode: 400,
      headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
      body: JSON.stringify({ error: `URL_${type.toUpperCase()} environment variable not set` }),
    };
  }

  // Serve from cache if fresh
  const now = Date.now();
  if (cache[type] && now - cache[type].ts < CACHE_TTL) {
    return {
      statusCode: 200,
      headers: {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*',
        'X-Cache': 'HIT',
      },
      body: JSON.stringify(cache[type].data),
    };
  }

  try {
    const buf    = await fetchExcel(SHARE_URLS[type]);
    const parsed = type === 'tenders' ? parseTenders(buf) : parseKrassat(buf);

    cache[type] = { ts: now, data: parsed };

    return {
      statusCode: 200,
      headers: {
        'Content-Type'                : 'application/json',
        'Access-Control-Allow-Origin' : '*',
        'X-Cache'                     : 'MISS',
        'X-Records'                   : String(Array.isArray(parsed) ? parsed.length : parsed.reports?.length ?? 0),
      },
      body: JSON.stringify(parsed),
    };
  } catch (err) {
    console.error('proxy error:', err);
    return {
      statusCode: 500,
      headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
      body: JSON.stringify({ error: err.message }),
    };
  }
};
