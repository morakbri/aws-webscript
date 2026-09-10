
// ==UserScript==
// @name         6az Position Grabber
// @namespace    https://amazon.sharepoint.com/
// @version      2.8.4
// @updateURL    https://github.com/morakbri/aws-webscript/raw/refs/heads/main/6az-position-grabber.user.js
// @downloadURL  https://github.com/morakbri/aws-webscript/raw/refs/heads/main/6az-position-grabber.user.js
// @description  Auto-extracts rack positions from the Excel file you're currently viewing for 6az — filters by AZ column, pastes values only, trusts sheet formatting. Groups by site with 2-row gaps.
// @author       @morakbri @tngujona
// @match        https://*.sharepoint.com/*
// @match        https://amazon.sharepoint.com/*
// @match        https://*.officeapps.live.com/*
// @match        https://view.officeapps.live.com/*
// @grant        GM_setClipboard
// @grant        GM_xmlhttpRequest
// @connect      *.sharepoint.com
// @require      https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js
// @run-at       document-idle
// ==/UserScript==

(function () {
    'use strict';
    if (window.top !== window.self) return;
    if (document.getElementById('az6-fab')) return;

    const TARGET_AZ = '6AZ';
    const VALID_RACK_PREFIXES = ['EC2', 'MV.BONSAI', 'BONSAI', 'EBS'];
    const ALERT_LOST_RES = ['lost reservation', 'lost res'];
    const ALERT_REJECTED = ['rejected', 'reject'];
    const SITE_GAP_ROWS = 2;
    const DESTINATION_DEFAULTS = {
        brickHandoffDate:'', dateWorkloadReceived:'',
        copper:'Not Installed', copperSignature:'Tech Alias',
        fiber:'Not Installed', fiberSignature:'Tech Alias',
        mnSignature:'Tech Alias',
        optics:'Check for Optics', brickPatching:'Tech Alias',
        sitePOC:'', positionStatus:''
    };

    // === STYLES ===
    const styles = document.createElement('style');
    styles.textContent = `
#az6-fab{position:fixed;bottom:30px;right:30px;width:56px;height:56px;border-radius:50%;background:#2563eb;color:#fff;font-size:14px;font-weight:700;border:none;cursor:pointer;z-index:2147483647;box-shadow:0 4px 14px rgba(37,99,235,.4);display:flex;align-items:center;justify-content:center;transition:transform .2s,box-shadow .2s;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif}
#az6-fab:hover{transform:scale(1.1);box-shadow:0 6px 20px rgba(37,99,235,.6)}
#az6-fab.az6-loading{animation:az6-pulse 1.2s infinite;background:#1d4ed8}
@keyframes az6-pulse{0%,100%{box-shadow:0 4px 14px rgba(37,99,235,.4)}50%{box-shadow:0 4px 24px rgba(37,99,235,.8)}}
#az6-popup-overlay{position:fixed;top:0;left:0;width:100vw;height:100vh;background:rgba(0,0,0,.6);z-index:2147483646;display:none;align-items:center;justify-content:center}
#az6-popup-overlay.az6-visible{display:flex}
#az6-popup{background:#111827;border:1px solid #374151;border-radius:12px;width:960px;max-width:94vw;max-height:88vh;overflow-y:auto;padding:24px;color:#e5e7eb;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;font-size:14px;box-shadow:0 25px 60px rgba(0,0,0,.5)}
#az6-popup h2{margin:0 0 4px;font-size:20px;color:#fff}
#az6-popup .az6-subtitle{color:#9ca3af;font-size:12px;margin-bottom:16px}
#az6-file-info{background:#1f2937;border:1px solid #374151;border-radius:6px;padding:10px 14px;margin-bottom:16px;display:flex;align-items:center;gap:10px}
#az6-file-info .az6-file-icon{font-size:20px}
#az6-file-info .az6-file-details{flex:1}
#az6-file-info .az6-file-name{color:#34d399;font-weight:600;font-size:13px}
#az6-file-info .az6-file-source{color:#6b7280;font-size:11px;margin-top:2px}
.az6-controls{display:flex;gap:12px;align-items:center;margin-bottom:16px;flex-wrap:wrap}
.az6-controls select{background:#1f2937;color:#e5e7eb;border:1px solid #374151;border-radius:6px;padding:8px 12px;font-size:13px;min-width:180px;cursor:pointer}
.az6-controls select:focus{outline:none;border-color:#2563eb}
.az6-btn{padding:8px 16px;border-radius:6px;border:none;font-size:13px;font-weight:600;cursor:pointer;transition:background .2s}
.az6-btn-primary{background:#2563eb;color:#fff}.az6-btn-primary:hover{background:#1d4ed8}.az6-btn-primary:disabled{background:#374151;color:#6b7280;cursor:not-allowed}
.az6-btn-success{background:#059669;color:#fff}.az6-btn-success:hover{background:#047857}
.az6-btn-close{background:#374151;color:#e5e7eb}.az6-btn-close:hover{background:#4b5563}
.az6-btn-refresh{background:#374151;color:#e5e7eb;font-size:12px;padding:6px 12px}.az6-btn-refresh:hover{background:#4b5563}
.az6-btn-optics{background:#0891b2;color:#fff}.az6-btn-optics:hover{background:#0e7490}
#az6-status{font-size:13px;margin-bottom:12px;min-height:20px;padding:6px 0}
#az6-status.az6-success{color:#34d399}#az6-status.az6-error{color:#f87171}#az6-status.az6-info{color:#60a5fa}
.az6-alert-banner{border-radius:8px;padding:12px 16px;margin-bottom:12px;font-size:13px;display:none}
.az6-alert-lost{background:#78350f;border:1px solid #d97706;color:#fde68a}
.az6-alert-rejected{background:#7f1d1d;border:1px solid #dc2626;color:#fca5a5}
.az6-alert-banner strong{display:block;margin-bottom:4px;font-size:14px}
.az6-alert-banner ul{margin:4px 0 0 16px;padding:0}.az6-alert-banner li{margin:2px 0}
#az6-results-table{width:100%;border-collapse:collapse;margin-bottom:16px;font-size:12px}
#az6-results-table th{background:#1f2937;color:#9ca3af;font-weight:600;text-align:left;padding:8px 10px;border-bottom:1px solid #374151;position:sticky;top:0}
#az6-results-table td{padding:6px 10px;border-bottom:1px solid #1f2937;color:#d1d5db}
#az6-results-table tr:hover td{background:#1f2937}
#az6-results-table tr.az6-row-lost td{background:#422006;color:#fde68a}
#az6-results-table tr.az6-row-rejected td{background:#450a0a;color:#fca5a5;text-decoration:line-through}
#az6-results-table tr.az6-site-divider td{background:#0f172a;padding:2px;border-bottom:2px solid #374151}
#az6-table-container{max-height:300px;overflow-y:auto;border:1px solid #374151;border-radius:6px;margin-bottom:16px;display:none}
#az6-output-box{background:#064e3b;border:1px solid #065f46;border-radius:6px;padding:12px;font-family:'Consolas','Monaco',monospace;font-size:12px;color:#6ee7b7;white-space:pre;overflow-x:auto;max-height:180px;overflow-y:auto;margin-bottom:12px;display:none}
#az6-copy-row{display:none;gap:12px;align-items:center;margin-bottom:8px;flex-wrap:wrap}
#az6-stats{font-size:12px;color:#9ca3af;margin-bottom:12px;display:none}
.az6-header-row{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:16px}
.az6-copy-hint{font-size:11px;color:#6b7280;width:100%;margin-top:4px}
#az6-optics-copy-row{display:none;gap:12px;align-items:center;margin-bottom:8px;flex-wrap:wrap;padding:10px 12px;background:#0c4a6e;border:1px solid #0369a1;border-radius:6px}
#az6-optics-copy-row .az6-optics-hint{font-size:11px;color:#7dd3fc;width:100%;margin-top:4px}
    `;
    document.head.appendChild(styles);

    // === UI ===
    const fab = document.createElement('button');
    fab.id = 'az6-fab'; fab.textContent = '6az';
    fab.title = '6az Position Grabber — Click to grab from current sheet';
    document.body.appendChild(fab);

    const overlay = document.createElement('div');
    overlay.id = 'az6-popup-overlay';
    overlay.innerHTML = `<div id="az6-popup">
        <div class="az6-header-row"><div><h2>6az Position Grabber</h2>
        <div class="az6-subtitle">Auto-extract rack positions — grouped by site, 2-row gaps — values only — v2.8.4</div></div>
        <button class="az6-btn az6-btn-close" id="az6-close-btn">Close</button></div>
        <div id="az6-file-info" style="display:none"><div class="az6-file-icon">📊</div>
        <div class="az6-file-details"><div class="az6-file-name" id="az6-detected-name">—</div>
        <div class="az6-file-source" id="az6-detected-source">—</div></div>
        <button class="az6-btn az6-btn-refresh" id="az6-refresh-btn">↻ Re-fetch</button></div>
        <div class="az6-controls"><select id="az6-sheet-select" disabled><option value="">-- Select Sheet --</option></select>
        <button class="az6-btn az6-btn-primary" id="az6-grab-btn" disabled>Grab 6az Positions</button></div>
        <div id="az6-status"></div>
        <div class="az6-alert-banner az6-alert-lost" id="az6-alert-lost"><strong>⚠️ LOST RESERVATIONS — Need to Re-Reserve!</strong><ul id="az6-lost-list"></ul></div>
        <div class="az6-alert-banner az6-alert-rejected" id="az6-alert-rejected"><strong>🚫 REJECTED POSITIONS</strong><ul id="az6-rejected-list"></ul></div>
        <div id="az6-stats"></div>
        <div id="az6-table-container"><table id="az6-results-table"><thead><tr>
        <th>Site</th><th>Position</th><th>Uplink Config</th><th>Asset</th><th>Rack Type</th><th>Rack Land Date</th><th>Brick</th><th>Notes</th>
        </tr></thead><tbody id="az6-results-body"></tbody></table></div>
        <div id="az6-output-box"></div>
        <div id="az6-copy-row"><button class="az6-btn az6-btn-success" id="az6-copy-btn">📋 Copy (Clean)</button>
        <button class="az6-btn az6-btn-primary" id="az6-copy-all-btn">📋 Copy (With Alerts)</button>
        <span id="az6-copy-status" style="font-size:12px;color:#34d399"></span>
        <div class="az6-copy-hint">Paste into row below header (Col A). Grouped by site with 2-row gaps. Values only — sheet formatting applies automatically. Cols T/U/V untouched.</div></div>
        <div id="az6-optics-copy-row"><button class="az6-btn az6-btn-optics" id="az6-copy-optics-btn">🔬 Copy Optics (Clean)</button>
        <span id="az6-optics-copy-status" style="font-size:12px;color:#7dd3fc"></span>
        <div class="az6-optics-hint">6-col optics TSV (Site · Position · Uplink Config · Asset · Brick · Land Date). Paste into Col A of the optics workload. Blank row between sites.</div></div>
    </div>`;
    document.body.appendChild(overlay);

    // === REFS ===
    const fileInfoEl=document.getElementById('az6-file-info'), detectedNameEl=document.getElementById('az6-detected-name'),
        detectedSourceEl=document.getElementById('az6-detected-source'), refreshBtn=document.getElementById('az6-refresh-btn'),
        sheetSelect=document.getElementById('az6-sheet-select'), grabBtn=document.getElementById('az6-grab-btn'),
        statusEl=document.getElementById('az6-status'), statsEl=document.getElementById('az6-stats'),
        alertLostEl=document.getElementById('az6-alert-lost'), lostListEl=document.getElementById('az6-lost-list'),
        alertRejectedEl=document.getElementById('az6-alert-rejected'), rejectedListEl=document.getElementById('az6-rejected-list'),
        tableContainer=document.getElementById('az6-table-container'), resultsBody=document.getElementById('az6-results-body'),
        outputBox=document.getElementById('az6-output-box'), copyRow=document.getElementById('az6-copy-row'),
        copyBtn=document.getElementById('az6-copy-btn'), copyAllBtn=document.getElementById('az6-copy-all-btn'),
        copyStatus=document.getElementById('az6-copy-status'), closeBtn=document.getElementById('az6-close-btn'),
        opticsCopyRow=document.getElementById('az6-optics-copy-row'), copyOpticsBtn=document.getElementById('az6-copy-optics-btn'),
        opticsCopyStatus=document.getElementById('az6-optics-copy-status');

    let workbook = null, currentFileUrl = null, lastGrabbedRows = [];

    // === HELPERS ===
    function isTarget6az(v) { return v && v.toUpperCase().trim() === TARGET_AZ; }
    function isValidRackType(rt) { if (!rt) return false; const u=rt.toUpperCase().trim(); return VALID_RACK_PREFIXES.some(p=>u.startsWith(p)); }
    function isExcelFile(url) { return /\.(xlsx|xlsm|xls)/i.test(url); }
    function getFileNameFromPath(url) { try { const p=decodeURIComponent(url.split('?')[0]).split('/'); return p[p.length-1]||null; } catch(e){ return null; } }

    function setStatus(msg, type) {
        statusEl.textContent = msg; statusEl.className = '';
        if (type==='success') statusEl.classList.add('az6-success');
        else if (type==='error') statusEl.classList.add('az6-error');
        else if (type==='info') statusEl.classList.add('az6-info');
    }
    function esc(str) { const d=document.createElement('div'); d.textContent=str; return d.innerHTML; }

    function sortBySite(rows) {
        return rows.sort((a,b)=>{
            const sA=a.site.toUpperCase(), sB=b.site.toUpperCase();
            const pA=sA.match(/(\d+|\D+)/g)||[], pB=sB.match(/(\d+|\D+)/g)||[];
            for (let i=0;i<Math.max(pA.length,pB.length);i++){
                const x=pA[i]||'', y=pB[i]||'', nA=parseInt(x,10), nB=parseInt(y,10);
                if (!isNaN(nA)&&!isNaN(nB)){ if(nA!==nB) return nA-nB; }
                else { if(x<y) return -1; if(x>y) return 1; }
            }
            return a.position.localeCompare(b.position,undefined,{numeric:true});
        });
    }
    function groupBySite(rows) {
        const groups=[], map=new Map();
        rows.forEach(r=>{ const k=r.site.toUpperCase().trim();
            if(!map.has(k)){ const g={site:r.site,rows:[]}; map.set(k,g); groups.push(g); }
            map.get(k).rows.push(r); });
        return groups;
    }

    // === v2.8.3 SAFE GUID EXTRACTION ===
    function cleanSourceDocGuid(raw) {
        if (!raw) return '';
        let s = raw;
        for (let i=0;i<2;i++){ try{ const d=decodeURIComponent(s); if(d===s) break; s=d; }catch(e){break;} }
        return s.replace(/[{}\s]/g, '').trim();
    }
    function formatGuid(guid) {
        if (!guid) return '';
        const c = guid.replace(/[^a-fA-F0-9]/g, '');
        if (c.length !== 32) return guid;
        return c.substring(0,8)+'-'+c.substring(8,12)+'-'+c.substring(12,16)+'-'+c.substring(16,20)+'-'+c.substring(20,32);
    }

    // === FILE URL DETECTION ===
    function detectExcelFileUrl() {
        const pageUrl=window.location.href, params=new URLSearchParams(window.location.search);
        const sourcedoc=params.get('sourcedoc'), fileParam=params.get('file');

        if (sourcedoc && pageUrl.includes('/_layouts/')) {
            const sitePath = extractSitePath(pageUrl);
            const cleanGuid = formatGuid(cleanSourceDocGuid(sourcedoc));
            console.log('[6az] sourcedoc raw:', sourcedoc, '→ cleanGuid:', cleanGuid);
            const wopiSrc = findWopiSrc(params, pageUrl);
            return { url:buildApiUrl(sitePath,cleanGuid), type:'sourcedoc', guid:cleanGuid,
                     fileName:fileParam?decodeURIComponent(fileParam):null, sitePath, wopiSrc };
        }
        if (pageUrl.includes('officeapps.live.com')) {
            const src=params.get('src')||params.get('wdOrigin');
            if (src&&isExcelFile(src)) return {url:src,type:'officeapps-src',fileName:getFileNameFromPath(src)};
            const wopiSrc=params.get('WOPISrc')||params.get('wopisrc');
            if (wopiSrc) return {url:wopiSrc+'/contents',type:'wopi',fileName:fileParam};
        }
        if (/\.(xlsx|xlsm|xls)(\?|$|#)/i.test(pageUrl)) { const u=pageUrl.split('?')[0]; return {url:u,type:'direct',fileName:getFileNameFromPath(u)}; }
        const idParam=params.get('id');
        if (idParam&&isExcelFile(idParam)) { const fp=idParam.startsWith('/')?idParam:'/'+idParam; return {url:window.location.origin+fp,type:'id-param',fileName:getFileNameFromPath(idParam)}; }
        return null;
    }
    function findWopiSrc(params, pageUrl) {
        const d=params.get('WOPISrc')||params.get('wopisrc'); if(d) return d;
        const m=pageUrl.match(/(https?:\/\/[^&\s]+_vti_bin\/wopi\.ashx\/files\/[a-f0-9]+)/i);
        return m?m[1]:null;
    }
    function extractSitePath(url) {
        try{ let p=new URL(url).pathname; p=p.replace(/^\/:x:\/[rs]\//,'/');
            const m=p.match(/^(\/(?:sites|teams)\/[^/]+)/i); if(m) return m[1]; }catch(e){}
        return '';
    }
    function buildApiUrl(sitePath, guid) {
        const base=sitePath?window.location.origin+sitePath:window.location.origin;
        return base+`/_api/web/GetFileById('${guid}')/$value`;
    }

    // === RESPONSE GUARD ===
    function validateBinaryResponse(buf) {
        if (!buf||buf.byteLength===0) return {ok:false,reason:'Empty response (0 bytes).'};
        const u=new Uint8Array(buf);
        if ((u[0]===0x50&&u[1]===0x4B&&u[2]===0x03&&u[3]===0x04)||(u[0]===0xD0&&u[1]===0xCF&&u[2]===0x11&&u[3]===0xE0)) return {ok:true};
        let txt='';
        try{ txt=new TextDecoder('utf-8',{fatal:false}).decode(u.slice(0,2000)).trim(); }catch(e){ return {ok:false,reason:'Unknown binary format.'}; }
        const lw=txt.toLowerCase();
        if (lw.includes('<!doctype')||lw.includes('<html')||lw.includes('<head')){
            if (lw.includes('sign in')||lw.includes('login')||lw.includes('idpinitiatedsignon')||lw.includes('microsoftonline'))
                return {ok:false,reason:'SharePoint returned a login page. Refresh and sign in again.'};
            if (lw.includes('access denied')||lw.includes('403')||lw.includes('unauthorized'))
                return {ok:false,reason:'SharePoint returned Access Denied.'};
            if (lw.includes('404')||lw.includes('not found'))
                return {ok:false,reason:'SharePoint returned 404 Not Found.'};
            return {ok:false,reason:'SharePoint returned HTML instead of the Excel file.'};
        }
        if (lw.startsWith('{')&&(lw.includes('"error"')||lw.includes('"odata.error"'))){
            try{ const e=JSON.parse(txt); return {ok:false,reason:'SP API error: '+(e?.['odata.error']?.message?.value||e?.error?.message?.value||'Unknown')}; }
            catch(e){ return {ok:false,reason:'SP API returned JSON error.'}; }
        }
        if (txt.length>0&&/^[\x20-\x7E\r\n\t]+$/.test(txt.substring(0,200)))
            return {ok:false,reason:'Response is plain text: '+txt.substring(0,200)};
        return {ok:true};
    }

    // === AUTH ===
    const SP_AUTH_HEADERS = {'Accept':'application/octet-stream','X-FORMS_BASED_AUTH_ACCEPTED':'f','Cache-Control':'no-cache'};

    async function fetchRequestDigest(sitePath) {
        const base=(sitePath?window.location.origin+sitePath:window.location.origin);
        try{ const r=await fetch(base+'/_api/contextinfo',{method:'POST',credentials:'include',
            headers:{'Accept':'application/json;odata=verbose','Content-Length':'0','X-FORMS_BASED_AUTH_ACCEPTED':'f'}});
            if(!r.ok)return null; const j=await r.json(); return j?.d?.GetContextWebInformation?.FormDigestValue||null;
        }catch(e){ console.log('[6az] digest fail:',e.message); return null; }
    }

    // === FETCH STRATEGIES ===
    function buildFetchStrategies(detected) {
        const o=window.location.origin, s=[];
        if (detected.guid&&detected.sitePath)
            s.push({label:'download.aspx',url:o+detected.sitePath+`/_layouts/15/download.aspx?UniqueId=${detected.guid}`,needsDigest:false});
        if (detected.guid&&detected.sitePath)
            s.push({label:'GetFileById',url:buildApiUrl(detected.sitePath,detected.guid),needsDigest:true});
        if (detected.wopiSrc)
            s.push({label:'WOPI-contents',url:detected.wopiSrc+'/contents',needsDigest:false});
        if (detected.guid&&detected.sitePath&&!detected.wopiSrc){
            const flat=detected.guid.replace(/-/g,'').toLowerCase();
            s.push({label:'WOPI-constructed',url:o+detected.sitePath+'/_vti_bin/wopi.ashx/files/'+flat+'/contents',needsDigest:false});
        }
        if (detected.type!=='sourcedoc'&&detected.url)
            s.push({label:'direct-url',url:detected.url,needsDigest:false});
        return s;
    }

    // === MULTI-STRATEGY FETCH ===
    async function fetchExcelFile(detected) {
        const strategies=buildFetchStrategies(detected); let digest=null, lastError='No strategies available.';
        for (const strat of strategies){
            console.log('[6az] Trying:',strat.label,'→',strat.url);
            const headers={...SP_AUTH_HEADERS};
            if (strat.needsDigest&&!digest&&detected.sitePath){ digest=await fetchRequestDigest(detected.sitePath); if(digest) console.log('[6az] Got digest.'); }
            if (strat.needsDigest&&digest) headers['X-RequestDigest']=digest;
            if (strat.url.includes('/_api/')) headers['binaryStringResponseBody']='true';

            try{ const r=await fetch(strat.url,{method:'GET',credentials:'include',headers,redirect:'follow'});
                if(r.ok){ const buf=await r.arrayBuffer(); const c=validateBinaryResponse(buf);
                    if(c.ok){ console.log('[6az] ✓ fetch()+',strat.label); return new Uint8Array(buf); }
                    console.warn('[6az] fetch()+',strat.label,'→',c.reason); lastError=c.reason;
                } else { console.log('[6az] fetch()+',strat.label,'→ HTTP',r.status); lastError='HTTP '+r.status+' from '+strat.label; }
            }catch(e){ console.log('[6az] fetch()+',strat.label,'→',e.message); lastError=e.message; }

            try{ const gm=await new Promise((res,rej)=>{
                GM_xmlhttpRequest({method:'GET',url:strat.url,responseType:'arraybuffer',headers,
                    onload(r){ if(r.status>=200&&r.status<300){ const c=validateBinaryResponse(r.response); c.ok?res(new Uint8Array(r.response)):rej(new Error(c.reason)); } else rej(new Error('HTTP '+r.status)); },
                    onerror(){ rej(new Error('Network error')); }});});
                console.log('[6az] ✓ GM+',strat.label); return gm;
            }catch(e){ console.log('[6az] GM+',strat.label,'→',e.message); lastError=e.message; }
        }
        throw new Error(lastError);
    }

    // === MAIN WORKFLOW ===
    async function autoGrab() {
        resetResults(); setStatus('Detecting Excel file...','info'); fab.classList.add('az6-loading');
        const detected=detectExcelFileUrl();
        if(!detected){ fab.classList.remove('az6-loading'); setStatus('No Excel file detected. Open an .xlsx/.xlsm in SharePoint first.','error'); overlay.classList.add('az6-visible'); return; }
        currentFileUrl=detected;
        const fileName=detected.fileName||getFileNameFromPath(detected.url)||'Unknown File';
        overlay.classList.add('az6-visible'); fileInfoEl.style.display='flex';
        detectedNameEl.textContent=fileName;
        detectedSourceEl.textContent=detected.sitePath?'Site: '+detected.sitePath+' | GUID: '+detected.guid+' | Method: '+detected.type:'Method: '+detected.type;
        setStatus('Fetching: '+fileName+'...','info');
        try{
            const data=await fetchExcelFile(detected);
            workbook=XLSX.read(data,{type:'array'});
            sheetSelect.innerHTML='<option value="">-- Select Sheet --</option>';
            workbook.SheetNames.forEach(n=>{ const o=document.createElement('option'); o.value=n; o.textContent=n; sheetSelect.appendChild(o); });
            sheetSelect.disabled=false; grabBtn.disabled=false;
            if(workbook.SheetNames.length===1) sheetSelect.value=workbook.SheetNames[0];
            fab.classList.remove('az6-loading');
            setStatus('File loaded! '+workbook.SheetNames.length+' sheet(s). Select a sheet and click Grab.','success');
            if(workbook.SheetNames.length===1) processSheet();
        }catch(err){ fab.classList.remove('az6-loading'); setStatus('Error fetching file: '+err.message,'error'); console.error('[6az] Fetch error:',err); }
    }
    function resetResults(){
        resultsBody.innerHTML=''; tableContainer.style.display='none'; outputBox.style.display='none'; outputBox.textContent='';
        copyRow.style.display='none'; opticsCopyRow.style.display='none'; statsEl.style.display='none';
        alertLostEl.style.display='none'; alertRejectedEl.style.display='none'; lostListEl.innerHTML=''; rejectedListEl.innerHTML=''; lastGrabbedRows=[];
    }

    // === PROCESS SHEET ===
    function processSheet(){
        if(!workbook){setStatus('No file loaded.','error');return;}
        const sheetName=sheetSelect.value; if(!sheetName){setStatus('Please select a sheet.','error');return;}
        setStatus('Processing...','info');
        try{
            const sheet=workbook.Sheets[sheetName], rows=XLSX.utils.sheet_to_json(sheet,{defval:''});
            if(!rows.length){setStatus('No data in sheet.','error');return;}
            const headers=Object.keys(rows[0]), colMap=findColumns(headers);
            if(!colMap.az){setStatus('Could not find "AZ" column.','error');return;}
            if(!colMap.site){setStatus('Could not find "Site" column.','error');return;}
            if(!colMap.position){setStatus('Could not find "Position" column.','error');return;}
            let grabbed=[], lostRes=[], rejPos=[], skipAz=0, skipRt=0;
            rows.forEach(row=>{
                const az=String(row[colMap.az]||'').trim(), site=String(row[colMap.site]||'').trim(), pos=String(row[colMap.position]||'').trim();
                if(!az||!pos) return; if(!isTarget6az(az)){skipAz++;return;}
                const rt=String(row[colMap.rackType]||'').trim(); if(rt&&!isValidRackType(rt)){skipRt++;return;}
                const notes=String(row[colMap.notes]||'').trim(), nl=notes.toLowerCase();
                const metroVal = colMap.metronomeLength ? String(row[colMap.metronomeLength]||'').trim() : '';
                const entry={site,position:pos,uplinkConfig:String(row[colMap.uplinkConfig]||'').trim(),
                    asset:String(row[colMap.asset]||'').trim(),rackType:rt,landDate:formatDate(row[colMap.landDate]),
                    brick:String(row[colMap.brick]||'').trim(),notes,
                    metronome: metroVal !== '' ? 'TRUE' : 'FALSE'};
                if(ALERT_LOST_RES.some(t=>nl.includes(t))){entry.status='lost';lostRes.push(entry);grabbed.push(entry);return;}
                if(ALERT_REJECTED.some(t=>nl.includes(t))){entry.status='rejected';rejPos.push(entry);grabbed.push(entry);return;}
                entry.status='ok'; grabbed.push(entry);
            });
            grabbed=sortBySite(grabbed); lastGrabbedRows=grabbed;
            displayResults(grabbed,lostRes,rejPos,skipAz,skipRt);
        }catch(err){setStatus('Error processing: '+err.message,'error');}
    }
    function findColumns(headers){
        const m={az:null,site:null,position:null,uplinkConfig:null,asset:null,rackType:null,landDate:null,brick:null,notes:null,metronomeLength:null};
        headers.forEach(h=>{ const l=h.toLowerCase().trim();
            if(l==='az'&&!m.az) m.az=h; else if(l==='site'&&!m.site) m.site=h;
            else if(l==='position'&&!m.position) m.position=h; else if(l.includes('uplink')&&!m.uplinkConfig) m.uplinkConfig=h;
            else if(l==='asset'&&!m.asset) m.asset=h; else if((l==='rack type'||l==='racktype')&&!m.rackType) m.rackType=h;
            else if((l.includes('rack land')||l==='land date'||l.includes('scheduled land'))&&!m.landDate) m.landDate=h;
            else if(l==='brick'&&!m.brick) m.brick=h; else if((l==='notes'||l==='note')&&!m.notes) m.notes=h;
            else if((l==='metronome length'||l==='metronome_length'||l==='metro length'||l==='metronome')&&!m.metronomeLength) m.metronomeLength=h;
        }); return m;
    }
    function formatDate(val){
        if(!val) return '';
        if(typeof val==='number'){ try{ const d=XLSX.SSF.parse_date_code(val); if(d) return String(d.m).padStart(2,'0')+'/'+String(d.d).padStart(2,'0')+'/'+d.y; }catch(e){return String(val);} }
        return String(val).trim();
    }

    // === TSV BUILDERS ===
    function buildRowTSV(e){
        const d=DESTINATION_DEFAULTS;
        return [e.site,e.position,e.uplinkConfig,e.asset,e.rackType,e.landDate,e.brick,
            d.brickHandoffDate,d.dateWorkloadReceived,d.copper,d.copperSignature,d.fiber,d.fiberSignature,
            e.metronome,d.mnSignature,d.optics,d.brickPatching,d.sitePOC,d.positionStatus].join('\t');
    }
    function buildEmptyRowTSV(){ return new Array(19).fill('').join('\t'); }
    function buildOpticsRowTSV(e){ return [e.site,e.position,e.uplinkConfig,e.asset,e.brick,e.landDate].join('\t'); }
    function buildEmptyOpticsRowTSV(){ return new Array(6).fill('').join('\t'); }

    function buildGroupedTSV(rows, includeAlerts){
        const t=includeAlerts?rows:rows.filter(r=>r.status==='ok'); if(!t.length) return '';
        const groups=groupBySite(t), parts=[];
        groups.forEach((g,i)=>{ g.rows.forEach(r=>parts.push(buildRowTSV(r)));
            if(i<groups.length-1) for(let j=0;j<SITE_GAP_ROWS;j++) parts.push(buildEmptyRowTSV()); });
        return parts.join('\n');
    }
    function buildOpticsTSV(rows){
        const clean=rows.filter(r=>r.status==='ok'); if(!clean.length) return '';
        const groups=groupBySite(clean), parts=[];
        groups.forEach((g,i)=>{ g.rows.forEach(r=>parts.push(buildOpticsRowTSV(r)));
            if(i<groups.length-1) parts.push(buildEmptyOpticsRowTSV()); });
        return parts.join('\n');
    }

    // === CLIPBOARD ===
    async function writeToClipboard(text){
        if(navigator.clipboard&&navigator.clipboard.writeText){ try{await navigator.clipboard.writeText(text);return true;}catch(e){} }
        try{ const ta=document.createElement('textarea'); ta.style.cssText='position:fixed;left:-9999px;top:-9999px;opacity:0;'; ta.value=text;
            document.body.appendChild(ta); ta.select(); const ok=document.execCommand('copy'); document.body.removeChild(ta); if(ok) return true; }catch(e){}
        try{ GM_setClipboard(text,'text'); return true; }catch(e){ return false; }
    }
    async function copyToClipboard(rows, includeAlerts){
        const tsv=buildGroupedTSV(rows,includeAlerts);
        if(!tsv){copyStatus.textContent='No rows to copy.';setTimeout(()=>{copyStatus.textContent='';},3000);return false;}
        return writeToClipboard(tsv);
    }

    // === DISPLAY RESULTS ===
    function displayResults(grabbed, lostRes, rejPos, skipAz, skipRt){
        const cleanCount=grabbed.filter(r=>r.status==='ok').length, siteGroups=groupBySite(grabbed);
        statsEl.style.display='block';
        statsEl.textContent=cleanCount+' positions grabbed | '+siteGroups.length+' sites | '+lostRes.length+' lost reservations | '+rejPos.length+' rejected | '+skipRt+' non-EC2/Bonsai/EBS skipped | '+skipAz+' non-6az skipped';

        if(lostRes.length>0){ alertLostEl.style.display='block'; lostListEl.innerHTML='';
            lostRes.forEach(r=>{ const li=document.createElement('li'); li.textContent=r.site+' — '+r.position+' ('+r.rackType+', Asset: '+r.asset+', Brick: '+r.brick+')'; lostListEl.appendChild(li); }); }
        if(rejPos.length>0){ alertRejectedEl.style.display='block'; rejectedListEl.innerHTML='';
            rejPos.forEach(r=>{ const li=document.createElement('li'); li.textContent=r.site+' — '+r.position+' ('+r.rackType+', Asset: '+r.asset+') — '+r.notes; rejectedListEl.appendChild(li); }); }
        if(!grabbed.length){setStatus('No matching 6az EC2/Bonsai/EBS positions found.','info');return;}

        setStatus(grabbed.length+' total across '+siteGroups.length+' sites. '+
            (lostRes.length>0?'⚠️ '+lostRes.length+' need re-reservation! ':'')+
            (rejPos.length>0?'🚫 '+rejPos.length+' rejected. ':''),
            lostRes.length>0||rejPos.length>0?'error':'success');

        resultsBody.innerHTML='';
        siteGroups.forEach((group,gIdx)=>{
            group.rows.forEach(r=>{
                const tr=document.createElement('tr');
                if(r.status==='lost') tr.className='az6-row-lost'; else if(r.status==='rejected') tr.className='az6-row-rejected';
                tr.innerHTML='<td>'+esc(r.site)+'</td><td>'+esc(r.position)+'</td><td>'+esc(r.uplinkConfig)+'</td><td>'+esc(r.asset)+'</td><td>'+esc(r.rackType)+'</td><td>'+esc(r.landDate)+'</td><td>'+esc(r.brick)+'</td><td>'+esc(r.notes)+'</td>';
                resultsBody.appendChild(tr);
            });
            if(gIdx<siteGroups.length-1){ const d=document.createElement('tr'); d.className='az6-site-divider'; d.innerHTML='<td colspan="8"></td>'; resultsBody.appendChild(d); }
        });
        tableContainer.style.display='block';

        const pLines=[], pGroups=siteGroups.slice(0,2);
        pGroups.forEach((g,i)=>{ g.rows.slice(0,2).forEach(r=>pLines.push(buildRowTSV(r)));
            if(g.rows.length>2) pLines.push('  ... +'+(g.rows.length-2)+' more from '+g.site);
            if(i<pGroups.length-1){pLines.push('');pLines.push('  ── 2-row gap ──');pLines.push('');} });
        if(siteGroups.length>2) pLines.push('\n... +'+(siteGroups.length-2)+' more site groups');
        outputBox.textContent='19-col TSV (A-S) — Grouped by site, 2-row gaps:\n\n'+pLines.join('\n');
        outputBox.style.display='block'; copyRow.style.display='flex';
        if(cleanCount>0) opticsCopyRow.style.display='flex';
    }

    // === EVENT HANDLERS ===
    fab.addEventListener('click',()=>{ if(overlay.classList.contains('az6-visible')) overlay.classList.remove('az6-visible'); else autoGrab(); });
    closeBtn.addEventListener('click',()=>overlay.classList.remove('az6-visible'));
    overlay.addEventListener('click',e=>{ if(e.target===overlay) overlay.classList.remove('az6-visible'); });
    document.addEventListener('keydown',e=>{ if(e.key==='Escape'&&overlay.classList.contains('az6-visible')) overlay.classList.remove('az6-visible'); });
    refreshBtn.addEventListener('click',()=>autoGrab());
    grabBtn.addEventListener('click',processSheet);

    copyBtn.addEventListener('click',async()=>{
        if(!lastGrabbedRows.length) return; copyStatus.textContent='Copying...';
        const ok=await copyToClipboard(lastGrabbedRows,false);
        copyStatus.textContent=ok?'✓ Copied! Grouped by site with 2-row gaps. Paste into Col A.':'⚠ Copy may have failed — try Ctrl+V anyway.';
        setTimeout(()=>{copyStatus.textContent='';},4000);
    });
    copyAllBtn.addEventListener('click',async()=>{
        if(!lastGrabbedRows.length) return; copyStatus.textContent='Copying...';
        const ok=await copyToClipboard(lastGrabbedRows,true);
        copyStatus.textContent=ok?'✓ Copied all (with alerts) — grouped, 2-row gaps. Paste into Col A.':'⚠ Copy may have failed — try Ctrl+V anyway.';
        setTimeout(()=>{copyStatus.textContent='';},4000);
    });
    copyOpticsBtn.addEventListener('click',async()=>{
        if(!lastGrabbedRows.length) return; opticsCopyStatus.textContent='Copying...';
        const tsv=buildOpticsTSV(lastGrabbedRows);
        if(!tsv){opticsCopyStatus.textContent='No clean rows to copy.';setTimeout(()=>{opticsCopyStatus.textContent='';},3000);return;}
        const ok=await writeToClipboard(tsv);
        opticsCopyStatus.textContent=ok?'✓ Optics copied! Paste into Col A of the optics workload.':'⚠ Copy may have failed — try Ctrl+V anyway.';
        setTimeout(()=>{opticsCopyStatus.textContent='';},4000);
    });

})();

