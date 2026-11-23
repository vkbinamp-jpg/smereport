/* Vanilla JavaScript only. Source Excel files are read-only. Patient rows stay in memory. */
(() => {
  "use strict";
  const cfg = window.APP_CONFIG || {};
  const $ = id => document.getElementById(id);
  const state = {
    rows: [], allRows: [], files: [], processedFiles: 0, issues: [], logs: [],
    page: 1, pageSize: 10, selectedCompany: null, bar: null, pie: null
  };
  const REQUIRED = ["uhid", "company"];
  const aliases = {
    uhid: ["uhid","hospital id","hospitalid","patient id","patientid","uhid no","uhid number","reg no","registration no"],
    company: ["ref","referred name","referredname","company","company name","organization","organisation","corporate","client name"],
    patientName: ["patient name","name","patient","member name"],
    date: ["checkup date","check up date","visit date","appointment date","date of visit","registration date","date"],
    gender: ["gender","sex"], age: ["age","patient age"]
  };
  const norm = s => String(s ?? "").trim().toLowerCase().replace(/[\s_\-./()]+/g," ").replace(/\s+/g," ").trim();
  const clean = s => String(s ?? "").replace(/\s+/g," ").trim();
  const escapeHtml = s => String(s ?? "").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
  const number = n => Number(n || 0).toLocaleString("en-IN");
  const log = (message, ok=true) => {
    state.logs.unshift({message, ok, time:new Date().toLocaleString()});
    renderLogs();
  };
  function status(message, type="") {
    $("statusMessage").textContent = message;
    $("statusMessage").className = "status-message " + type;
  }

  // This is intentionally a simple client-side page gate, NOT secure authentication.
  $("loginForm").addEventListener("submit", e => {
    e.preventDefault();
    if ($("username").value.trim() === "admin" && $("password").value === "admin") {
      $("loginScreen").classList.add("hidden");
      $("app").classList.remove("hidden");
      $("loginError").textContent = "";
      $("driveStatusText").textContent = "Public Drive folder mode";
      if (publicDriveReady()) loadDriveFiles();
      else status("Configure config.js with the Google API Key and fixed Folder ID to load data.", "error");
    } else $("loginError").textContent = "Invalid username or password.";
  });
  $("logoutBtn").addEventListener("click", () => {
    state.rows=[]; state.allRows=[]; state.files=[]; state.processedFiles=0;
    $("app").classList.add("hidden"); $("loginScreen").classList.remove("hidden");
    $("password").value=""; renderAll();
    $("driveStatusText").textContent="Public Drive folder not loaded";
    $("refreshBtn").disabled=true;
  });

  function publicDriveReady() {
    return cfg.GOOGLE_API_KEY && !cfg.GOOGLE_API_KEY.startsWith("YOUR_") &&
      cfg.DRIVE_FOLDER_ID && !cfg.DRIVE_FOLDER_ID.startsWith("YOUR_");
  }
  async function connectDrive() {
    if (!publicDriveReady()) {
      status("Setup required: copy config.js.example to config.js and enter Google API Key and fixed Folder ID.", "error");
      return;
    }
    await loadDriveFiles();
  }
  // Public folder mode: no Google OAuth sign-in. Only public Viewer files are readable.
  $("refreshBtn").addEventListener("click", loadDriveFiles);

  async function driveFetch(url) {
    const response = await fetch(url);
    if (!response.ok) {
      const body=await response.text().catch(()=> "");
      throw new Error("Google Drive API error ("+response.status+"): "+body.slice(0,180));
    }
    return response;
  }
  async function loadDriveFiles() {
    if (!publicDriveReady()) { connectDrive(); return; }
    status("Reading Excel file list from the fixed Google Drive folder...");
    $("refreshBtn").disabled=true;
    state.rows=[]; state.allRows=[]; state.files=[]; state.processedFiles=0; state.issues=[];
    try {
      const q = `'${cfg.DRIVE_FOLDER_ID}' in parents and trashed = false`;
      const params = new URLSearchParams({q, pageSize:"1000", fields:"files(id,name,mimeType,modifiedTime,size),nextPageToken", orderBy:"name"});
      const res = await driveFetch("https://www.googleapis.com/drive/v3/files?"+params.toString()+"&key="+encodeURIComponent(cfg.GOOGLE_API_KEY));
      const data=await res.json();
      const excelFiles=(data.files||[]).filter(f => /\.(xlsx|xls)$/i.test(f.name) &&
        f.mimeType !== "application/vnd.google-apps.spreadsheet");
      state.files=excelFiles;
      $("fileCount").textContent=excelFiles.length+" files";
      log("Found "+excelFiles.length+" Excel file(s) in configured folder");
      if (!excelFiles.length) {
        status("No .xlsx or .xls files found in the configured folder.", "error");
        renderAll(); return;
      }
      let processed=0;
      for (const file of excelFiles) {
        try {
          status(`Processing ${++processed} of ${excelFiles.length}: ${file.name}`);
          const fileRes=await driveFetch(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(file.id)}?alt=media&key=${encodeURIComponent(cfg.GOOGLE_API_KEY)}`);
          const buffer=await fileRes.arrayBuffer();
          const workbook=XLSX.read(buffer,{type:"array",cellDates:true});
          let sheetCount=0;
          for (const sheetName of workbook.SheetNames) {
            const sheet=workbook.Sheets[sheetName];
            const matrix=XLSX.utils.sheet_to_json(sheet,{header:1,defval:"",raw:false});
            const result=parseWorksheet(matrix,file.name,sheetName);
            if (result.error) {
              state.issues.push({file:file.name,sheet:sheetName,issue:result.error,rows:"—"});
              continue;
            }
            state.rows.push(...result.rows);
            state.issues.push(...result.issues);
            sheetCount++;
          }
          if (sheetCount > 0) state.processedFiles++;
          log(`${file.name}: processed ${sheetCount} worksheet(s)`);
        } catch(err) {
          state.issues.push({file:file.name,sheet:"—",issue:"File could not be processed: "+err.message,rows:"—"});
          log(file.name+": processing failed ("+err.message+")",false);
        }
      }
      state.allRows=state.rows.slice();
      populateFilterOptions();
      setDefaultDates();
      applyFilters();
      $("lastSync").textContent="Last sync: "+new Date().toLocaleString();
      status(`Completed. Read ${state.rows.length.toLocaleString("en-IN")} patient rows from ${excelFiles.length} Excel file(s).`, "success");
    } catch(err) {
      status(err.message,"error"); log("Drive read failed: "+err.message,false);
    } finally {
      $("refreshBtn").disabled=!publicDriveReady();
    }
  }

  function findHeaderRow(matrix) {
    let best={index:-1,map:{},score:0};
    const limit=Math.min(matrix.length,20);
    for(let i=0;i<limit;i++){
      const cells=(matrix[i]||[]).map(norm);
      const map={};
      for(const [field,list] of Object.entries(aliases)){
        const idx=cells.findIndex(c=>list.includes(c));
        if(idx>=0) map[field]=idx;
      }
      const score=Object.keys(map).length+(map.uhid!==undefined?3:0)+(map.company!==undefined?3:0);
      if(score>best.score) best={index:i,map,score};
    }
    return best.index>=0?best:null;
  }
  function parseWorksheet(matrix,file,sheet) {
    const header=findHeaderRow(matrix);
    if(!header || header.map.uhid===undefined || header.map.company===undefined)
      return {error:"Required columns UHID and/or Ref (Company) not detected. Manual mapping is required."};
    const rows=[], issues=[];
    let missingUhid=0,missingCompany=0,invalidDates=0,blank=0;
    for(let i=header.index+1;i<matrix.length;i++){
      const r=matrix[i]||[];
      if(r.every(v=>clean(v)==="")){blank++;continue;}
      const get=key=>header.map[key]===undefined?"":r[header.map[key]];
      const uhid=clean(get("uhid"));
      const company=clean(get("company"));
      const date=parseDate(get("date"));
      if(!uhid) missingUhid++;
      if(!company) missingCompany++;
      if(get("date")!=="" && !date) invalidDates++;
      rows.push({
        uhid, uhidKey:uhid.toLowerCase(),
        company, companyKey:company.toLowerCase().replace(/\s+/g," "),
        patientName:clean(get("patientName")),
        date, dateRaw:clean(get("date")),
        gender:clean(get("gender")), age:clean(get("age")),
        file, sheet, excelRow:i+1
      });
    }
    if(missingUhid) issues.push({file,sheet,issue:"Rows with missing UHID",rows:missingUhid});
    if(missingCompany) issues.push({file,sheet,issue:"Rows with missing Ref / company",rows:missingCompany});
    if(invalidDates) issues.push({file,sheet,issue:"Rows with invalid checkup date",rows:invalidDates});
    return {rows,issues};
  }
  function parseDate(value) {
    if(value instanceof Date && !isNaN(value)) return dateOnly(value);
    const s=clean(value); if(!s) return null;
    // Excel serial date (when parser returns a number)
    if(typeof value==="number" && value>20000 && value<80000){
      const d=XLSX.SSF.parse_date_code(value); if(d) return `${d.y}-${String(d.m).padStart(2,"0")}-${String(d.d).padStart(2,"0")}`;
    }
    let m=s.match(/^(\d{1,2})[-/. ]([A-Za-z]{3,9})[-/. ,](\d{2,4})$/);
    if(m){const months=["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];const mo=months.findIndex(x=>m[2].toLowerCase().startsWith(x));if(mo>=0){let y=+m[3];if(y<100)y+=y<50?2000:1900;return `${y}-${String(mo+1).padStart(2,"0")}-${String(+m[1]).padStart(2,"0")}`;}}
    m=s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
    if(m){let y=+m[3];if(y<100)y+=y<50?2000:1900;const a=+m[1],b=+m[2]; // Indian date convention: DD/MM/YYYY
      if(a>=1&&a<=31&&b>=1&&b<=12)return `${y}-${String(b).padStart(2,"0")}-${String(a).padStart(2,"0")}`;
    }
    const d=new Date(s); if(!isNaN(d)) return dateOnly(d);
    return null;
  }
  function dateOnly(d){return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;}
  function setDefaultDates(){
    const dates=state.rows.map(r=>r.date).filter(Boolean).sort();
    if(dates.length){$("fromDate").value=dates[0];$("toDate").value=dates[dates.length-1];}
    else {$("fromDate").value="";$("toDate").value="";}
  }
  function populateFilterOptions(){
    const companies=[...new Map(state.rows.filter(r=>r.company&&r.companyKey).map(r=>[r.companyKey,r.company])).values()].sort((a,b)=>a.localeCompare(b));
    $("companyFilter").innerHTML='<option value="">All companies</option>'+companies.map(c=>`<option value="${escapeHtml(c.toLowerCase().replace(/\s+/g," "))}">${escapeHtml(c)}</option>`).join("");
    $("fileFilter").innerHTML='<option value="">All files</option>'+state.files.map(f=>`<option value="${escapeHtml(f.name)}">${escapeHtml(f.name)}</option>`).join("");
  }
  function filteredRows(){
    const from=$("fromDate").value,to=$("toDate").value,company=$("companyFilter").value,uhid=$("uhidSearch").value.trim().toLowerCase(),file=$("fileFilter").value;
    return state.allRows.filter(r=>(!from||(r.date && r.date>=from))&&(!to||(r.date && r.date<=to))&&
      (!company||r.companyKey===company)&&(!uhid||r.uhidKey.includes(uhid))&&(!file||r.file===file));
  }
  function applyFilters(){
    state.rows=filteredRows(); state.page=1; state.selectedCompany=null;
    $("patientDetails").classList.add("hidden"); renderAll();
  }
  $("applyFiltersBtn").addEventListener("click",applyFilters);
  $("clearFiltersBtn").addEventListener("click",()=>{
    $("companyFilter").value="";$("uhidSearch").value="";$("fileFilter").value="";
    setDefaultDates();applyFilters();
  });
  $("exportBtn").addEventListener("click",exportCsv);
  $("prevPage").addEventListener("click",()=>{if(state.page>1){state.page--;renderCompanyTable();}});
  $("nextPage").addEventListener("click",()=>{state.page++;renderCompanyTable();});
  $("closePatient").addEventListener("click",()=>$("patientDetails").classList.add("hidden"));
  $("clearLogsBtn").addEventListener("click",()=>{state.logs=[];renderLogs();});

  function summarize(rows){
    const map=new Map(); let missingUhid=0,missingCompany=0,duplicates=0,invalidDates=0;
    const seen=new Set();
    for(const r of rows){
      if(!r.uhid)missingUhid++;
      if(!r.company)missingCompany++;
      if(r.dateRaw&&!r.date)invalidDates++;
      if(!r.uhid||!r.company)continue;
      const key=r.companyKey+"||"+r.uhidKey;
      if(seen.has(key))duplicates++; else seen.add(key);
      if(!map.has(r.companyKey))map.set(r.companyKey,{name:r.company,uhids:new Set(),records:0,patients:new Map()});
      const c=map.get(r.companyKey); c.records++;
      c.uhids.add(r.uhidKey);
      if(!c.patients.has(r.uhidKey))c.patients.set(r.uhidKey,{uhid:r.uhid,name:r.patientName,date:r.date,records:0});
      const p=c.patients.get(r.uhidKey);p.records++;
      if(!p.name&&r.patientName)p.name=r.patientName;
      if(r.date&&(!p.date||r.date<p.date))p.date=r.date;
    }
    const companies=[...map.values()].map(c=>({...c,unique:c.uhids.size})).sort((a,b)=>b.unique-a.unique||a.name.localeCompare(b.name));
    return {companies,uniqueTotal:companies.reduce((n,c)=>n+c.unique,0),recordTotal:rows.length,missingUhid,missingCompany,duplicates,invalidDates};
  }
  function renderAll(){
    const s=summarize(state.rows);
    $("kpiCompanies").textContent=number(s.companies.length);
    $("kpiPatients").textContent=number(s.uniqueTotal);
    $("kpiRecords").textContent=number(s.recordTotal);
    $("kpiFiles").textContent=number(state.processedFiles);
    $("kpiMissingUhid").textContent=number(s.missingUhid);
    $("kpiMissingCompany").textContent=number(s.missingCompany);
    $("qualityDuplicates").textContent=number(s.duplicates);
    $("qualityDates").textContent=number(s.invalidDates);
    $("qualityUhid").textContent=number(s.missingUhid);
    $("qualityCompany").textContent=number(s.missingCompany);
    renderCompanyTable(s);renderCharts(s);renderQuality();renderLogs();
    $("exportBtn").disabled=!state.rows.length;
  }
  function renderCompanyTable(summary=summarize(state.rows)){
    const all=summary.companies,total=summary.uniqueTotal,pages=Math.max(1,Math.ceil(all.length/state.pageSize));
    if(state.page>pages)state.page=pages;
    const shown=all.slice((state.page-1)*state.pageSize,state.page*state.pageSize);
    $("companyTable").innerHTML=shown.length?shown.map((c,i)=>`<tr><td>${(state.page-1)*state.pageSize+i+1}</td><td><a class="company-link" href="#" data-company="${escapeHtml(c.name)}">${escapeHtml(c.name)}</a></td><td><strong>${number(c.unique)}</strong></td><td>${number(c.records)}</td><td>${total?(c.unique/total*100).toFixed(1):"0.0"}%</td></tr>`).join(""):'<tr><td colspan="5" class="empty">No matching records.</td></tr>';
    $("companyTableFoot").innerHTML=all.length?`<tr><td colspan="2">Grand Total</td><td>${number(total)}</td><td>${number(summary.recordTotal)}</td><td>100%</td></tr>`:"";
    $("tableInfo").textContent=all.length+" companies";
    $("pageInfo").textContent=`Page ${state.page} of ${pages}`;
    $("prevPage").disabled=state.page<=1;$("nextPage").disabled=state.page>=pages;
    $("companyTable").querySelectorAll("[data-company]").forEach(a=>a.addEventListener("click",e=>{e.preventDefault();showPatients(a.dataset.company);}));
  }
  function showPatients(companyName){
    const s=summarize(state.rows),c=s.companies.find(x=>x.name===companyName);
    if(!c)return;
    $("patientHeading").textContent=companyName;
    $("patientSubheading").textContent=`${number(c.unique)} unique patients • ${number(c.records)} records`;
    const patients=[...c.patients.values()].sort((a,b)=>a.uhid.localeCompare(b.uhid,undefined,{numeric:true}));
    $("patientTable").innerHTML=patients.map(p=>`<tr><td>${escapeHtml(p.uhid)}</td><td>${escapeHtml(p.name||"—")}</td><td>${escapeHtml(p.date||"—")}</td><td>${number(p.records)}</td></tr>`).join("");
    $("patientDetails").classList.remove("hidden");$("patientDetails").scrollIntoView({behavior:"smooth",block:"start"});
  }
  function renderCharts(s){
    const top=s.companies.slice(0,10);
    if(state.bar)state.bar.destroy();if(state.pie)state.pie.destroy();
    if(!window.Chart)return;
    state.bar=new Chart($("barChart"),{type:"bar",data:{labels:top.map(c=>c.name),datasets:[{label:"Unique Patients",data:top.map(c=>c.unique),backgroundColor:"#2878b9",borderRadius:5}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:"y",plugins:{legend:{display:false}},scales:{x:{beginAtZero:true,ticks:{precision:0}},y:{ticks:{autoSkip:false}}}}});
    state.pie=new Chart($("pieChart"),{type:"doughnut",data:{labels:top.map(c=>c.name),datasets:[{data:top.map(c=>c.unique),backgroundColor:["#1769aa","#ef7d22","#20a078","#7d65bd","#e0b23e","#5d9ec9","#db6b72","#6a9d52","#9b7a65","#7f8c9a"],borderWidth:2,borderColor:"#fff"}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:"bottom",labels:{boxWidth:10,font:{size:10}}}}}});
  }
  function renderQuality(){
    const issues=state.issues;
    $("qualityTable").innerHTML=issues.length?issues.map(x=>`<tr><td>${escapeHtml(x.file)}</td><td>${escapeHtml(x.sheet)}</td><td>${escapeHtml(x.issue)}</td><td>${escapeHtml(x.rows)}</td></tr>`).join(""):'<tr><td colspan="4" class="empty">No validation findings.</td></tr>';
  }
  function renderLogs(){
    $("logs").innerHTML=state.logs.length?state.logs.map(x=>`<div class="log-row"><span><strong class="${x.ok?"log-ok":"log-error"}">${x.ok?"SUCCESS":"ERROR"}</strong> &nbsp;${escapeHtml(x.message)}</span><time class="muted">${escapeHtml(x.time)}</time></div>`).join(""):'<p class="empty">No processing activity yet.</p>';
  }
  function exportCsv(){
    const s=summarize(state.rows);
    const lines=[["S.No.","Company Name","Unique Patients","Total Records","Contribution"]];
    s.companies.forEach((c,i)=>lines.push([i+1,c.name,c.unique,c.records,(s.uniqueTotal?c.unique/s.uniqueTotal*100:0).toFixed(2)+"%"]));
    lines.push(["","Grand Total",s.uniqueTotal,s.recordTotal,"100%"]);
    const csv=lines.map(row=>row.map(v=>'"'+String(v??"").replace(/"/g,'""')+'"').join(",")).join("\r\n");
    const blob=new Blob(["\ufeff"+csv],{type:"text/csv;charset=utf-8"});
    const url=URL.createObjectURL(blob),a=document.createElement("a");a.href=url;a.download="Company_Wise_Health_Checkup_MIS.csv";a.click();URL.revokeObjectURL(url);
  }
  renderAll();
})();