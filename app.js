/* Vanilla JavaScript only. Source Excel files are read-only. Patient rows stay in memory. */
(() => {
  "use strict";
  const cfg = window.APP_CONFIG || {};
  const $ = id => document.getElementById(id);
  const state = {
    rows: [], allRows: [], files: [], hospitalRows: [], allHospitalRows: [], mode: "sme", processedFiles: 0, issues: [], logs: [],
    page: 1, pageSize: 10, selectedCompany: null, bar: null, pie: null
  };
  const REQUIRED = ["uhid", "company"];
  const aliases = {
    uhid: ["uhid","hospital id","hospitalid","patient id","patientid","uhid no","uhid number","reg no","registration no"],
    company: ["ref","referred name","referredname","company","company name","referrer name","referrername","organization","organisation","corporate","client name"],
    patientName: ["patient name","name","patient","member name"],
    date: ["checkup date","check up date","visit date","appointment date","date of visit","registration date","invoice date","item billed date","date"],
    gender: ["gender","sex"], age: ["age","patient age"], item: ["item name","itemname"], department: ["department name","department"], billed: ["total billed","totalbilled"], paid: ["net paid","netpaid"], due: ["due amount","dueamount"]
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
    status("Reading CSV / Excel file list from the fixed Google Drive folder...");
    $("refreshBtn").disabled=true;
    state.rows=[]; state.allRows=[]; state.hospitalRows=[]; state.allHospitalRows=[]; state.files=[]; state.processedFiles=0; state.issues=[];
    try {
      const folderSpecs=[{id:cfg.DRIVE_FOLDER_ID,kind:"sme"},{id:cfg.HOSPITAL_FOLDER_ID,kind:"hospital"}].filter(x=>x.id);
      let allDriveFiles=[];
      for(const folder of folderSpecs){let pageToken="";do{const q=`\'${folder.id}\' in parents and trashed = false`;const params=new URLSearchParams({q,pageSize:"1000",fields:"files(id,name,mimeType,modifiedTime,size),nextPageToken",orderBy:"name"});if(pageToken)params.set("pageToken",pageToken);const res=await driveFetch("https://www.googleapis.com/drive/v3/files?"+params.toString()+"&key="+encodeURIComponent(cfg.GOOGLE_API_KEY));const data=await res.json();(data.files||[]).forEach(f=>allDriveFiles.push({...f,_kind:folder.kind}));pageToken=data.nextPageToken||"";}while(pageToken);}
      const dataFiles=allDriveFiles.filter(f => {
        const name=String(f.name||"");
        const mime=String(f.mimeType||"").toLowerCase();
        return /\.(csv|xlsx|xls)$/i.test(name) ||
          mime==="text/csv" || mime==="application/csv" || mime==="text/comma-separated-values";
      }).filter(f => !String(f.mimeType||"").startsWith("application/vnd.google-apps."));
      state.files=dataFiles;
      log("Drive folder listing: "+allDriveFiles.length+" item(s). "+(allDriveFiles.length ? allDriveFiles.map(f=>`${f.name} [${f.mimeType||"unknown"}]`).join(" | ") : "No direct child files returned."));
      $("fileCount").textContent=dataFiles.length+" files";
      log("Found "+dataFiles.length+" CSV / Excel file(s) in configured folder");
      if (!dataFiles.length) {
        const listing=allDriveFiles.length ? allDriveFiles.map(f=>`${f.name} (${f.mimeType||"unknown"})`).join(", ") : "Drive API returned no direct child files";
        status("No CSV / Excel files detected. Folder items returned: "+listing+". Confirm files are directly in this folder and have .csv, .xlsx or .xls extension.", "error");
        renderAll(); return;
      }
      let processed=0;
      for (const file of dataFiles) {
        try {
          status(`Processing ${++processed} of ${dataFiles.length}: ${file.name}`);
          const fileRes=await driveFetch(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(file.id)}?alt=media&key=${encodeURIComponent(cfg.GOOGLE_API_KEY)}`);
          const buffer=await fileRes.arrayBuffer();
          const isCsv=/\.csv$/i.test(file.name);
          // SheetJS detects CSV delimiters/encoding and Excel workbook formats from the downloaded bytes.
          const workbook=XLSX.read(buffer,{type:"array",cellDates:true, raw:false, codepage:65001});
          let sheetCount=0;
          for (const sheetName of workbook.SheetNames) {
            const sheet=workbook.Sheets[sheetName];
            const matrix=XLSX.utils.sheet_to_json(sheet,{header:1,defval:"",raw:false});
            const result=file._kind==="hospital"?parseHospitalWorksheet(matrix,file.name,sheetName):parseWorksheet(matrix,file.name,sheetName);
            if (result.error) {
              state.issues.push({file:file.name,sheet:sheetName,issue:result.error,rows:"—"});
              continue;
            }
            if(file._kind==="hospital")state.hospitalRows.push(...result.rows);else state.rows.push(...result.rows);
            state.issues.push(...(result.issues||[]));
            sheetCount++;
          }
          if (sheetCount > 0) state.processedFiles++;
          log(`${file.name}: processed ${sheetCount} ${isCsv?"CSV data set":"worksheet(s)"}`);
        } catch(err) {
          state.issues.push({file:file.name,sheet:"—",issue:"File could not be processed: "+err.message,rows:"—"});
          log(file.name+": processing failed ("+err.message+")",false);
        }
      }
      state.allRows=state.rows.slice();state.allHospitalRows=state.hospitalRows.slice();
      populateFilterOptions();
      setDefaultDates();
      applyFilters();
      $("lastSync").textContent="Last sync: "+new Date().toLocaleString();
      status(`Completed. Read ${state.rows.length.toLocaleString("en-IN")} patient rows from ${dataFiles.length} CSV / Excel file(s).`, "success");
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
      const get=key=>header.map[key]===undefined?"":r[header.map[key]];
      const uhid=clean(get("uhid"));
      const itemName=clean(get("item"));
      if (header.map.item !== undefined && !/^health/i.test(itemName)) continue;
      // Billing exports append totals/refunds after a blank separator. Stop at that boundary.
      if(!uhid && rows.length && r.every(v=>clean(v)==="")) break;
      if(r.every(v=>clean(v)==="")){blank++;continue;}
      // Ignore report footer/summary rows that are not patient-level records.
      if(!uhid && rows.length) continue;
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

  function parseHospitalWorksheet(matrix,file,sheet){
    let headerIndex=-1, map={};
    for(let i=0;i<Math.min(matrix.length,30);i++){
      const cells=(matrix[i]||[]).map(norm);
      const find=(names)=>cells.findIndex(c=>names.includes(c));
      const candidate={date:find(["invoice date","item billed date"]),department:find(["department name","department"]),billed:find(["total billed"]),paid:find(["net paid"]),due:find(["due amount"])};
      if(candidate.department>=0&&candidate.billed>=0&&candidate.paid>=0){headerIndex=i;map=candidate;break;}
    }
    if(headerIndex<0)return {error:"Hospital report columns not found (Department Name, Total Billed, Net Paid)."};
    const out=[]; let section="sales";
    for(let i=headerIndex+1;i<matrix.length;i++){
      const r=matrix[i]||[]; const first=norm(r[0]);
      if(first==="refunds"||first==="summary")break;
      if(r.every(v=>clean(v)===""))continue;
      const department=clean(r[map.department]); if(!department)continue;
      const billed=toAmount(r[map.billed]), paid=toAmount(r[map.paid]), due=map.due>=0?toAmount(r[map.due]):Math.max(0,billed-paid);
      const date=map.date>=0?parseDate(r[map.date]):null;
      out.push({department,date,billed,paid,due,file,sheet});
    }
    return {rows:out};
  }
  function toAmount(v){const n=Number(String(v??"").replace(/[,₹\s]/g,""));return Number.isFinite(n)?n:0;}
  function filteredHospitalRows(){const from=$("fromDate").value,to=$("toDate").value,file=$("fileFilter").value;return state.allHospitalRows.filter(r=>(!from||(r.date&&r.date>=from))&&(!to||(r.date&&r.date<=to))&&(!file||r.file===file));}
  function renderHospital(){
    const map=new Map(); for(const r of state.rows){const x=map.get(r.department)||{department:r.department,billed:0,paid:0,due:0,rows:0};x.billed+=r.billed;x.paid+=r.paid;x.due+=r.due;x.rows++;map.set(r.department,x);}
    const list=[...map.values()].sort((a,b)=>b.paid-a.paid); const totals=list.reduce((a,x)=>({billed:a.billed+x.billed,paid:a.paid+x.paid,due:a.due+x.due,rows:a.rows+x.rows}),{billed:0,paid:0,due:0,rows:0});
    $("kpiCompanies").parentElement.querySelector("p").textContent="Total Departments";$("kpiPatients").parentElement.querySelector("p").textContent="Net Paid Collection";$("kpiRecords").parentElement.querySelector("p").textContent="Total Billed";$("kpiFiles").parentElement.querySelector("p").textContent="Files Processed";$("kpiMissingUhid").parentElement.querySelector("p").textContent="Due Amount";$("kpiMissingCompany").parentElement.querySelector("p").textContent="Sales Rows";$("kpiCompanies").textContent=number(list.length);$("kpiPatients").textContent="₹"+number(totals.paid);$("kpiRecords").textContent="₹"+number(totals.billed);$("kpiFiles").textContent=number(state.processedFiles);$("kpiMissingUhid").textContent="₹"+number(totals.due);$("kpiMissingCompany").textContent=number(totals.rows);
    const body=$("companyTable"); body.innerHTML=list.length?list.map((x,i)=>`<tr><td>${i+1}</td><td>${escapeHtml(x.department)}</td><td>₹${number(x.billed)}</td><td>₹${number(x.paid)}</td><td>₹${number(x.due)}</td></tr>`).join(""):'<tr><td colspan="5" class="empty">No department collection data.</td></tr>';
    $("companyTable").parentElement.querySelector("thead").innerHTML="<tr><th>S.No.</th><th>Department</th><th>Total Billed</th><th>Net Paid</th><th>Due Amount</th></tr>";
    $("companyTableFoot").innerHTML=`<tr><td colspan="2">Grand Total</td><td>₹${number(totals.billed)}</td><td>₹${number(totals.paid)}</td><td>₹${number(totals.due)}</td></tr>`;$("tableInfo").textContent=list.length+" departments";$("pageInfo").textContent="Department summary";$("prevPage").disabled=true;$("nextPage").disabled=true;$("exportBtn").disabled=!list.length;
    if(state.bar)state.bar.destroy();if(state.pie)state.pie.destroy();if(window.Chart&&list.length){state.bar=new Chart($("barChart"),{type:"bar",data:{labels:list.slice(0,12).map(x=>x.department),datasets:[{label:"Net Paid",data:list.slice(0,12).map(x=>x.paid),backgroundColor:"#2878b9"}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:"y"}});}
  }
  function setMode(mode){state.mode=mode;$("smeModeBtn").classList.toggle("active",mode==="sme");$("hospitalModeBtn").classList.toggle("active",mode==="hospital");$("companyFilter").parentElement.style.display=mode==="sme"?"":"none";$("uhidSearch").parentElement.style.display=mode==="sme"?"":"none";$("companyReportTitle").textContent=mode==="sme"?"Company-wise Summary Report":"Department-wise Collection Report";applyFilters();}
  $("smeModeBtn").addEventListener("click",()=>setMode("sme"));$("hospitalModeBtn").addEventListener("click",()=>setMode("hospital"));
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
    const today=dateOnly(new Date());
    $("fromDate").value=today;
    $("toDate").value=today;
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
    state.rows=state.mode==="hospital"?filteredHospitalRows():filteredRows(); state.page=1; state.selectedCompany=null;
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
    if(state.mode==="hospital"){renderHospital();renderQuality();renderLogs();return;}
    ["Total Companies","Total Unique Patients","Total Health Checkup Records","Files Processed","Missing UHID","Missing Company"].forEach((t,i)=>$( ["kpiCompanies","kpiPatients","kpiRecords","kpiFiles","kpiMissingUhid","kpiMissingCompany"][i]).parentElement.querySelector("p").textContent=t);
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
    if(state.mode==="hospital"){const map=new Map();state.rows.forEach(r=>{const x=map.get(r.department)||{department:r.department,billed:0,paid:0,due:0};x.billed+=r.billed;x.paid+=r.paid;x.due+=r.due;map.set(r.department,x);});const lines=[["Department","Total Billed","Net Paid","Due Amount"],...[...map.values()].map(x=>[x.department,x.billed,x.paid,x.due])];downloadCsv(lines,"Hospital_Department_Collection.csv");return;}
    const s=summarize(state.rows);
    const lines=[["S.No.","Company Name","Unique Patients","Total Records","Contribution"]];
    s.companies.forEach((c,i)=>lines.push([i+1,c.name,c.unique,c.records,(s.uniqueTotal?c.unique/s.uniqueTotal*100:0).toFixed(2)+"%"]));
    lines.push(["","Grand Total",s.uniqueTotal,s.recordTotal,"100%"]);
    const csv=lines.map(row=>row.map(v=>'"'+String(v??"").replace(/"/g,'""')+'"').join(",")).join("\r\n");
    downloadCsv(lines,"Company_Wise_Health_Checkup_MIS.csv");
  }
  function downloadCsv(lines,name){const csv=lines.map(row=>row.map(v=>'"'+String(v??"").replace(/"/g,'""')+'"').join(",")).join("\r\n");const blob=new Blob(["\ufeff"+csv],{type:"text/csv;charset=utf-8"});const url=URL.createObjectURL(blob),a=document.createElement("a");a.href=url;a.download=name;a.click();URL.revokeObjectURL(url);}
  setDefaultDates();
  renderAll();
})();