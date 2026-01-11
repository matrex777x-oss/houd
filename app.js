/* app.js (V5) — Modern UI + flexible print options (portrait/landscape + choose report) */
(function(){
  const $ = (id)=>document.getElementById(id);
  const dayNamesAr = ["الأحد","الاثنين","الثلاثاء","الأربعاء","الخميس","الجمعة","السبت"];

  function pad2(n){ return (n<10? "0":"")+n; }
  function formatDate(d){ return pad2(d.getDate()) + "/" + pad2(d.getMonth()+1) + "/" + d.getFullYear(); }
  function ymd(d){ return d.getFullYear()+"-"+pad2(d.getMonth()+1)+"-"+pad2(d.getDate()); }
  function formatDateTime(d){ return formatDate(d)+" "+pad2(d.getHours())+":"+pad2(d.getMinutes())+":"+pad2(d.getSeconds()); }
  function formatTime12(d){
    let h = d.getHours(), m=d.getMinutes(), s=d.getSeconds();
    const am = h < 12;
    let h12 = h % 12; if(h12===0) h12=12;
    return pad2(h12)+":"+pad2(m)+":"+pad2(s)+" "+(am?"AM":"PM");
  }
  function minutesDiff(later, earlier){ return Math.round((later.getTime()-earlier.getTime())/60000); }
  function setStatus(msg, isErr=false){
    $("status").innerHTML = isErr ? "<span class='err'>"+msg+"</span>" : msg;
  }
  function toMidnight(d){ return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0,0,0,0); }
  function addDays(d, n){ const x = new Date(d.getTime()); x.setDate(x.getDate()+n); return x; }

  // ===== Parsing =====
  function parseTimeString(t){
    t = (t||"").toString().trim();
    let isPM = /م/.test(t);
    let isAM = /ص/.test(t);
    t = t.replace(/[صم]/g,"").trim();
    const parts = t.split(":").map(x=>x.trim()).filter(Boolean);
    if(parts.length<2) return null;
    let hh = parseInt(parts[0],10);
    let mm = parseInt(parts[1],10);
    let ss = parts.length>=3 ? parseInt(parts[2],10) : 0;
    if(Number.isNaN(hh)||Number.isNaN(mm)||Number.isNaN(ss)) return null;
    if(isPM && hh<12) hh += 12;
    if(isAM && hh===12) hh = 0;
    return {hh,mm,ss};
  }
  function parseDateString(ds){
    ds = (ds||"").toString().trim();
    const m = ds.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if(!m) return null;
    return {dd:+m[1], mm:+m[2], yy:+m[3]};
  }
  function parseDateTimeFlexible(datePart, timePart){
    if(timePart==null || String(timePart).trim()===""){
      const txt = String(datePart||"").trim().replace(/\s+/g," ");
      const m = txt.match(/^(\d{1,2}\/\d{1,2}\/\d{4})\s+(.+)$/);
      if(!m) return null;
      const d = parseDateString(m[1]); if(!d) return null;
      const t = parseTimeString(m[2]); if(!t) return null;
      return new Date(d.yy, d.mm-1, d.dd, t.hh, t.mm, t.ss, 0);
    } else {
      const d = parseDateString(String(datePart||"").trim()); if(!d) return null;
      const t = parseTimeString(String(timePart||"").trim()); if(!t) return null;
      return new Date(d.yy, d.mm-1, d.dd, t.hh, t.mm, t.ss, 0);
    }
  }
  function parseRaw(text){
    const lines = text.split(/\r?\n/).map(l=>l.trim()).filter(Boolean);
    const out = [];
    for(const line of lines){
      let parts = line.split("\t").map(x=>x.trim()).filter(Boolean);
      if(parts.length < 5) parts = line.split(/\s{2,}/).map(x=>x.trim()).filter(Boolean);
      if(parts.length < 5) continue;
      let empId, name, dept, datePart, timePart, deviceId;
      if(parts.length >= 6){
        empId = parseInt(parts[0],10);
        name = parts[1]||""; dept = parts[2]||"";
        datePart = parts[3]||""; timePart = parts[4]||"";
        deviceId = parts[5]||"";
      } else {
        empId = parseInt(parts[0],10);
        name = parts[1]||""; dept = parts[2]||"";
        datePart = parts[3]||""; timePart = "";
        deviceId = parts[4]||"";
      }
      if(!Number.isFinite(empId)) continue;
      const dt = parseDateTimeFlexible(datePart, timePart);
      if(!dt || isNaN(dt.getTime())) continue;
      out.push({empId, name, dept, deviceId, dateTime: dt});
    }
    return out;
  }
  function dedupPunches(punches, dedupSeconds){
    punches.sort((a,b)=> a.empId-b.empId || a.dateTime-b.dateTime);
    const out = [];
    const lastByEmp = new Map();
    for(const p of punches){
      const last = lastByEmp.get(p.empId);
      if(last){
        const diff = (p.dateTime.getTime()-last.dateTime.getTime())/1000;
        if(diff >= 0 && diff < dedupSeconds) continue;
      }
      out.push(p);
      lastByEmp.set(p.empId, p);
    }
    return out;
  }
  function computeShiftDate(dt, rollHour){
    const mid = toMidnight(dt);
    const threshold = new Date(mid.getTime());
    threshold.setHours(rollHour,0,0,0);
    if(dt < threshold){
      const prev = new Date(mid.getTime());
      prev.setDate(prev.getDate()-1);
      return prev;
    }
    return mid;
  }

  // ===== Schedule =====
  const schedule = {};
  function parseHHMM(s){
    const m = String(s||"").trim().match(/^(\d{1,2}):(\d{2})$/);
    if(!m) return null;
    const hh = +m[1], mm = +m[2];
    if(hh<0||hh>23||mm<0||mm>59) return null;
    return {hh,mm};
  }
  function hhmmOk(v){ return /^\d{1,2}:\d{2}$/.test(String(v||"")) && parseHHMM(v) !== null; }
  function defaultScheduleFor(empId){
    const isManager = (empId===1);
    const workAll = [true,true,true,true,true,true,true];
    if(empId===3 || empId===4){
      return {isManager, weekdayStart:"16:00", weekdayEnd:"01:00", weekendStart:"16:00", weekendEnd:"01:00", workDays:workAll};
    }
    if(empId===7){
      return {isManager, weekdayStart:"05:00", weekdayEnd:"16:00", weekendStart:"12:00", weekendEnd:"16:00", workDays:workAll};
    }
    return {isManager, weekdayStart:"14:30", weekdayEnd:"01:00", weekendStart:"14:30", weekendEnd:"01:00", workDays:workAll};
  }
  function ensureScheduleRows(employeeMap, maxId){
    for(let id=1; id<=maxId; id++){
      if(!schedule[id]){
        schedule[id] = { name: employeeMap.get(id) || "", ...defaultScheduleFor(id) };
      } else {
        schedule[id].name = schedule[id].name || employeeMap.get(id) || "";
        const def = defaultScheduleFor(id);
        schedule[id].isManager = (schedule[id].isManager ?? def.isManager);
        schedule[id].weekdayStart = schedule[id].weekdayStart || def.weekdayStart;
        schedule[id].weekdayEnd = schedule[id].weekdayEnd || def.weekdayEnd;
        schedule[id].weekendStart = schedule[id].weekendStart || def.weekendStart;
        schedule[id].weekendEnd = schedule[id].weekendEnd || def.weekendEnd;
        schedule[id].workDays = schedule[id].workDays || def.workDays;
      }
    }
  }
  function renderScheduleTable(maxId){
    const tbl = $("tblSchedule");
    const cols = ["EmployeeID","Name","Manager?","Weekday Start","Weekday End","Weekend Start","Weekend End","Sun","Mon","Tue","Wed","Thu","Fri","Sat"];
    tbl.innerHTML = "";
    const thead = document.createElement("thead");
    const trh = document.createElement("tr");
    for(const c of cols){ const th=document.createElement("th"); th.textContent=c; th.dataset.col=c; trh.appendChild(th); }
    thead.appendChild(trh); tbl.appendChild(thead);

    const tbody = document.createElement("tbody");
    for(let id=1; id<=maxId; id++){
      const s = schedule[id];
      const tr = document.createElement("tr");

      const tdId = document.createElement("td"); tdId.textContent=String(id); tr.appendChild(tdId);

      const tdName = document.createElement("td");
      const inName = document.createElement("input");
      inName.value = s.name || ""; inName.placeholder="name";
      inName.addEventListener("input", ()=>{ schedule[id].name = inName.value; });
      tdName.appendChild(inName); tr.appendChild(tdName);

      const tdMgr = document.createElement("td"); tdMgr.className="dayCell";
      const ckMgr = document.createElement("input"); ckMgr.type="checkbox"; ckMgr.checked=!!s.isManager;
      ckMgr.addEventListener("change", ()=>{ schedule[id].isManager = ckMgr.checked; });
      tdMgr.appendChild(ckMgr); tr.appendChild(tdMgr);

      tr.appendChild(tdTimeInput(id,"weekdayStart"));
      tr.appendChild(tdTimeInput(id,"weekdayEnd"));
      tr.appendChild(tdTimeInput(id,"weekendStart"));
      tr.appendChild(tdTimeInput(id,"weekendEnd"));

      for(let d=0; d<7; d++){
        const td = document.createElement("td"); td.className="dayCell";
        const ck = document.createElement("input"); ck.type="checkbox"; ck.checked=!!s.workDays[d];
        ck.addEventListener("change", ()=>{ schedule[id].workDays[d]=ck.checked; });
        td.appendChild(ck); tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }
    tbl.appendChild(tbody);

    function tdTimeInput(empId, field){
      const td = document.createElement("td");
      const inp = document.createElement("input");
      inp.value = schedule[empId][field]; inp.placeholder="HH:MM";
      inp.addEventListener("input", ()=>{ schedule[empId][field] = inp.value.trim(); });
      td.appendChild(inp); return td;
    }
  }
  function readScheduleFromUI(maxId){
    for(let id=1; id<=maxId; id++){
      const s = schedule[id];
      for(const f of ["weekdayStart","weekdayEnd","weekendStart","weekendEnd"]){
        if(!hhmmOk(s[f])){
          alert(`خطأ في جدول الدوام: EmployeeID=${id} field=${f} value=${s[f]}\nاكتب HH:MM مثل 14:30`);
          return false;
        }
      }
      if(!Array.isArray(s.workDays) || s.workDays.length!==7) s.workDays=[true,true,true,true,true,true,true];
    }
    return true;
  }
  function scheduleForDate(empId, shiftDate){
    const s = schedule[empId] || defaultScheduleFor(empId);
    const wd = shiftDate.getDay();
    const isWeekend = (wd===5 || wd===6);
    const startStr = isWeekend ? s.weekendStart : s.weekdayStart;
    const endStr   = isWeekend ? s.weekendEnd   : s.weekdayEnd;
    const st = parseHHMM(startStr), en = parseHHMM(endStr);
    if(!st || !en) return null;

    const start = new Date(shiftDate.getFullYear(), shiftDate.getMonth(), shiftDate.getDate(), st.hh, st.mm, 0, 0);
    let end = new Date(shiftDate.getFullYear(), shiftDate.getMonth(), shiftDate.getDate(), en.hh, en.mm, 0, 0);
    if(end <= start) end.setDate(end.getDate()+1);

    return {start, end, isWorkday: !!(s.workDays && s.workDays[wd]), isManager: !!s.isManager};
  }

  // ===== Build outputs =====
  function buildPunchRows(clean, rollHour){
    return clean.map(p=>{
      const sd = computeShiftDate(p.dateTime, rollHour);
      return {
        EmployeeID: p.empId,
        Name: p.name,
        Dept: p.dept,
        DeviceID: p.deviceId,
        DateTime: formatDateTime(p.dateTime),
        ShiftDate: formatDate(sd),
        DayName: dayNamesAr[sd.getDay()],
        Time12h: formatTime12(p.dateTime)
      };
    });
  }

  function buildDailyWithAbsences(clean, rollHour, graceMin, monthVal, maxId){
    const groups = new Map();
    let minSD=null, maxSD=null;

    for(const p of clean){
      const sd = computeShiftDate(p.dateTime, rollHour);
      if(!minSD || sd < minSD) minSD = sd;
      if(!maxSD || sd > maxSD) maxSD = sd;

      const key = p.empId + "|" + ymd(sd);
      let g = groups.get(key);
      if(!g){ g = {empId:p.empId, name:p.name, shiftDate:sd, punches:[]}; groups.set(key,g); }
      g.punches.push(p.dateTime);
      if(!g.name && p.name) g.name=p.name;
    }
    if(!minSD || !maxSD) return [];

    let rangeStart = toMidnight(minSD);
    let rangeEnd = toMidnight(maxSD);
    if(monthVal !== "ALL"){
      const [yyyy, mm] = monthVal.split("-").map(x=>parseInt(x,10));
      rangeStart = new Date(yyyy, mm-1, 1, 0,0,0,0);
      rangeEnd = new Date(yyyy, mm, 0, 0,0,0,0);
    }

    const daily = [];
    const presentSet = new Set();

    // present/irregular days
    for(const g of groups.values()){
      const key = g.empId + "|" + ymd(g.shiftDate);
      presentSet.add(key);

      if(monthVal !== "ALL"){
        const [yyyy, mm] = monthVal.split("-").map(x=>parseInt(x,10));
        if(g.shiftDate.getFullYear() !== yyyy || (g.shiftDate.getMonth()+1) !== mm) continue;
      }

      g.punches.sort((a,b)=>a-b);
      const sched = scheduleForDate(g.empId, g.shiftDate);
      if(!sched) continue;

      const {start, end, isManager} = sched;
      const punchCount = g.punches.length;
      const actualIn = g.punches[0];
      const actualOutRaw = g.punches[punchCount-1];
      const status = (punchCount===1) ? "IRREGULAR" : "REGULAR";

      let actualOut = actualOutRaw;
      if(actualOut < actualIn){ actualOut = new Date(actualOut.getTime()); actualOut.setDate(actualOut.getDate()+1); }
      const workMin = (punchCount===1) ? 0 : Math.max(0, minutesDiff(actualOut, actualIn));

      let lateMin = 0;
      if(!isManager){
        const rawLate = Math.max(0, minutesDiff(actualIn, start));
        lateMin = Math.max(0, rawLate - graceMin);
      }

      let otMin = 0;
      if(!isManager && punchCount>1){
        let outForOt = actualOutRaw;
        if(outForOt < start){ outForOt = new Date(outForOt.getTime()); outForOt.setDate(outForOt.getDate()+1); }
        otMin = Math.max(0, minutesDiff(outForOt, end));
      }

      daily.push({
        __rowType: status,
        EmployeeID: g.empId,
        Name: schedule[g.empId]?.name || g.name || "",
        ShiftDate: formatDate(g.shiftDate),
        DayName: dayNamesAr[g.shiftDate.getDay()],
        Status: status,
        SchedStart: formatTime12(start),
        SchedEnd: formatTime12(end),
        ActualIn: formatTime12(actualIn),
        ActualOut: (punchCount===1 ? "—" : formatTime12(actualOutRaw)),
        WorkMinutes: workMin,
        LateMinutes: lateMin,
        OTMinutes: otMin,
        Present: (isManager ? "" : "YES"),
        Absent: (isManager ? "" : "NO")
      });
    }

    // absences based on schedule days
    for(let empId=1; empId<=maxId; empId++){
      const s = schedule[empId] || defaultScheduleFor(empId);
      if(s.isManager) continue;
      for(let d = new Date(rangeStart.getTime()); d <= rangeEnd; d = addDays(d,1)){
        const sd = toMidnight(d);
        const sch = scheduleForDate(empId, sd);
        if(!sch || !sch.isWorkday) continue;
        const key = empId + "|" + ymd(sd);
        if(presentSet.has(key)) continue;

        daily.push({
          __rowType: "ABSENT",
          EmployeeID: empId,
          Name: s.name || "",
          ShiftDate: formatDate(sd),
          DayName: dayNamesAr[sd.getDay()],
          Status: "ABSENT",
          SchedStart: formatTime12(sch.start),
          SchedEnd: formatTime12(sch.end),
          ActualIn: "—",
          ActualOut: "—",
          WorkMinutes: 0,
          LateMinutes: 0,
          OTMinutes: 0,
          Present: "NO",
          Absent: "YES"
        });
      }
    }

    daily.sort((a,b)=>{
      const aId = (a.EmployeeID==="TOTAL"? 1e9 : +a.EmployeeID);
      const bId = (b.EmployeeID==="TOTAL"? 1e9 : +b.EmployeeID);
      if(aId!==bId) return aId-bId;
      return a.ShiftDate.localeCompare(b.ShiftDate);
    });
    return daily;
  }

  function buildSummary(daily){
    const byEmp = new Map();
    for(const r of daily){
      if(r.EmployeeID === "TOTAL") continue;
      const empId = +r.EmployeeID;
      const isManager = !!(schedule[empId]?.isManager);
      let s = byEmp.get(empId);
      if(!s){
        s = {EmployeeID: empId, Name: schedule[empId]?.name || r.Name || "",
             DaysPresent:0, DaysAbsent:0, IrregularDays:0, TotalLateMin:0, TotalOTMin:0, TotalWorkHours:0};
        byEmp.set(empId,s);
      }
      if(isManager) continue;
      if(r.Present==="YES") s.DaysPresent++;
      if(r.Absent==="YES") s.DaysAbsent++;
      if(r.Status==="IRREGULAR") s.IrregularDays++;
      s.TotalLateMin += (+r.LateMinutes||0);
      s.TotalOTMin += (+r.OTMinutes||0);
      s.TotalWorkHours += (+r.WorkMinutes||0)/60;
    }
    const out = Array.from(byEmp.values()).sort((a,b)=>a.EmployeeID-b.EmployeeID);
    for(const x of out) x.TotalWorkHours = Math.round(x.TotalWorkHours*100)/100;
    return out;
  }

  // Totals rows
  function addTotalsRowDaily(rows){
    if(!rows || rows.length===0) return rows;
    let work=0, late=0, ot=0, present=0, absent=0, irregular=0;
    for(const r of rows){
      work += (+r.WorkMinutes||0);
      late += (+r.LateMinutes||0);
      ot += (+r.OTMinutes||0);
      if(r.Present==="YES") present++;
      if(r.Absent==="YES") absent++;
      if(r.Status==="IRREGULAR") irregular++;
    }
    const total = {};
    Object.keys(rows[0]).forEach(k=> total[k]="");
    total.__rowType = "TOTAL";
    total.EmployeeID = "TOTAL";
    total.Name = "المجموع";
    total.Status = `Present=${present}, Absent=${absent}, Irregular=${irregular}`;
    total.WorkMinutes = work; total.LateMinutes = late; total.OTMinutes = ot;
    return rows.concat([total]);
  }
  function addTotalsRowSummary(rows){
    // Totals row in Monthly Summary: only aggregate meaningful totals across employees.
    // We keep: TotalWorkHours + TotalOTMin. Other "totals" are intentionally blank.
    if(!rows || rows.length===0) return rows;

    let ot=0, wh=0;
    for(const r of rows){
      ot += (+r.TotalOTMin||0);
      wh += (+r.TotalWorkHours||0);
    }

    const total = {};
    Object.keys(rows[0]).forEach(k=> total[k]="");
    total.__rowType = "TOTAL";
    total.EmployeeID = "TOTAL";
    total.Name = "المجموع";
    if("TotalWorkHours" in total) total.TotalWorkHours = Math.round(wh*100)/100;
    if("TotalOTMin" in total) total.TotalOTMin = Math.round(ot);

    return rows.concat([total]);
  }

  function renderTable(tbl, rows){
    tbl.innerHTML = "";
    if(!rows || rows.length===0){
      tbl.innerHTML = "<tr><td style='padding:14px;color:#9fb0d0'>لا توجد بيانات بعد.</td></tr>";
      return;
    }
    const cols = Object.keys(rows[0]).filter(c=>c!=="__rowType");
    const thead = document.createElement("thead");
    const trh = document.createElement("tr");
    for(const c of cols){ const th=document.createElement("th"); th.textContent=c; th.dataset.col=c; trh.appendChild(th); }
    thead.appendChild(trh); tbl.appendChild(thead);

    const tbody = document.createElement("tbody");
    for(const r of rows){
      const tr = document.createElement("tr");
      if(r.__rowType==="TOTAL") tr.classList.add("totalRow");
      if(r.__rowType==="ABSENT") tr.classList.add("absentRow");
      if(r.__rowType==="IRREGULAR") tr.classList.add("irregularRow");

      for(const c of cols){
        const td = document.createElement("td");
        td.dataset.col = c;
        td.textContent = (r[c]===null||r[c]===undefined) ? "" : String(r[c]);
        if(c==="ActualIn" || c==="ActualOut") td.classList.add("actualCell");
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }
    tbl.appendChild(tbody);
  }

  // TSV export
  function downloadTSV(filename, rows){
    if(!rows || rows.length===0){ alert("لا توجد بيانات للتصدير."); return; }
    const cols = Object.keys(rows[0]).filter(c=>c!=="__rowType");
    const esc = (v)=>{ v=(v===null||v===undefined)?"":String(v); return v.replace(/\t/g," ").replace(/\r?\n/g," "); };
    const tsv = [cols.map(esc).join("\t")].concat(rows.map(r=>cols.map(c=>esc(r[c])).join("\t"))).join("\n");
    const blob = new Blob([tsv], {type:"text/tab-separated-values;charset=utf-8"});
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(()=>{ URL.revokeObjectURL(a.href); a.remove(); }, 0);
  }

  // ===== XLSX export (Excel) =====
  // Uses SheetJS (xlsx) loaded from CDN in index.html.
  // If you are offline and XLSX is not loaded, you can still use TSV (Excel can open it).
  function ensureXLSX(){
    if(typeof XLSX === "undefined"){
      alert("XLSX library not loaded.\n- افتح الصفحة مع اتصال انترنت (لتحميل المكتبة)\n- أو استخدم TSV (يفتح في Excel).");
      return false;
    }
    return true;
  }

  function exportRowsToXLSX(rows, sheetName, fileName){
    if(!ensureXLSX()) return;
    if(!rows || rows.length===0){ alert("لا توجد بيانات للتصدير."); return; }
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, sheetName || "Sheet1");
    XLSX.writeFile(wb, fileName || "attendance.xlsx");
  }

  function exportDatasetXLSX(kind){
    // kind: "summary" | "daily" | "punches" | "employee"
    if(kind==="summary") return exportRowsToXLSX(addTotalsRowSummary(state.summary || []), "MonthlySummary", "Monthly_Summary.xlsx");
    if(kind==="daily")   return exportRowsToXLSX(addTotalsRowDaily(state.dailyFiltered || []), "DailyAttendance", "Daily_Attendance.xlsx");
    if(kind==="punches") return exportRowsToXLSX(state.punchesFiltered || [], "Punches", "Punches.xlsx");
    if(kind==="employee"){
      const empId = $("empSelect")?.value;
      const empName = (state.employeeMap?.get(String(empId))?.name) || "Employee";
      const rows = addTotalsRowDaily(state.empDailyRows || []);
      const safe = `Employee_${empId}_${empName}.xlsx`.replace(/[\\/:*?"<>|]/g,"_");
      return exportRowsToXLSX(rows, "EmployeeReport", safe);
    }
  }


  // Filters utils
  function computeMonthsFromDaily(daily){
    const set = new Map();
    for(const r of daily){
      const parts = r.ShiftDate.split("/");
      const key = parts[2]+"-"+parts[1];
      if(!set.has(key)) set.set(key, {key, label: parts[1]+"/"+parts[2]});
    }
    return Array.from(set.values()).sort((a,b)=>a.key.localeCompare(b.key));
  }
  function populateMonthFilter(months){
    const sel = $("monthFilter");
    const current = sel.value;
    sel.innerHTML = '<option value="ALL">كل البيانات</option>';
    for(const m of months){ const opt=document.createElement("option"); opt.value=m.key; opt.textContent=m.label; sel.appendChild(opt); }
    sel.value = (current && Array.from(sel.options).some(o=>o.value===current)) ? current : "ALL";
  }
  function buildEmployeeNameMap(raw){
    const m = new Map();
    for(const p of raw){
      if(!m.has(p.empId)) m.set(p.empId, p.name || "");
      else if(!m.get(p.empId) && p.name) m.set(p.empId, p.name);
    }
    return m;
  }
  function populateEmployeeFilter(maxId){
    const sel = $("employeeFilter");
    const current = sel.value;
    sel.innerHTML = '<option value="ALL">كل الموظفين</option>';
    for(let id=1; id<=maxId; id++){
      const opt=document.createElement("option");
      opt.value=String(id);
      const nm=schedule[id]?.name||"";
      opt.textContent=id + (nm?(" - "+nm):"");
      sel.appendChild(opt);
    }
    sel.value = (current && Array.from(sel.options).some(o=>o.value===current)) ? current : "ALL";
  }
  function computeBest(summaryRows){
    if(!summaryRows || summaryRows.length===0) return "—";
    const scored = summaryRows
      .filter(r => r.EmployeeID!=="TOTAL" && !schedule[r.EmployeeID]?.isManager)
      .map(r=>{
        const score = (r.DaysPresent*10) - (r.DaysAbsent*20) - ((r.TotalLateMin||0)/10) - ((r.IrregularDays||0)*2);
        return {id:r.EmployeeID, name:r.Name||"", score};
      });
    if(scored.length===0) return "—";
    scored.sort((a,b)=>b.score-a.score);
    return `${scored[0].id} - ${scored[0].name}`;
  }

  // Employee report
  function renderEmployeeReport(){
    const empVal = $("employeeFilter").value;
    const monthVal = $("monthFilter").value;
    const kpi1=$("empKpi"), kpi2=$("empKpi2");

    if(empVal==="ALL"){
      kpi1.style.display="none"; kpi2.style.display="none";
      renderTable($("tblEmpDaily"), []);
      renderTable($("tblEmpPunches"), []);
      return;
    }
    const empId = parseInt(empVal,10);
    const dailyAll = state.dailyFiltered.filter(r => r.EmployeeID===empId);
    const punchesAll = state.punchesFiltered.filter(r => r.EmployeeID===empId);

    let daysP=0, daysA=0, late=0, ot=0, workMin=0;
    for(const r of dailyAll){
      if(r.Present==="YES") daysP++;
      if(r.Absent==="YES") daysA++;
      late += (+r.LateMinutes||0);
      ot   += (+r.OTMinutes||0);
      workMin += (+r.WorkMinutes||0);
    }
    const workH = Math.round((workMin/60)*100)/100;

    let first="—", last="—";
    if(punchesAll.length>0){
      first = punchesAll[0].DateTime + " (" + punchesAll[0].Time12h + ")";
      last  = punchesAll[punchesAll.length-1].DateTime + " (" + punchesAll[punchesAll.length-1].Time12h + ")";
    }

    $("eDaysP").textContent=daysP;
    $("eDaysA").textContent=daysA;
    $("eLate").textContent=late;
    $("eOT").textContent=ot;
    $("eWorkH").textContent=workH;
    $("eFirst").textContent=first;
    $("eLast").textContent=last;
    kpi1.style.display=""; kpi2.style.display="";

    renderTable($("tblEmpDaily"), addTotalsRowDaily(dailyAll));
    renderTable($("tblEmpPunches"), punchesAll);

    $("btnExportEmpDaily").onclick = ()=> downloadTSV(`employee_${empId}_daily_${monthVal==="ALL"?"all":monthVal}.tsv`, addTotalsRowDaily(dailyAll));
    $("btnExportEmpPunches").onclick = ()=> downloadTSV(`employee_${empId}_punches_${monthVal==="ALL"?"all":monthVal}.tsv`, punchesAll);
  }

  // Tabs
  function setActiveTab(tabName){
    state.activeTab = tabName;
    document.querySelectorAll(".tab").forEach(x=>x.classList.remove("active"));
    document.querySelectorAll(".view").forEach(v=>v.style.display="none");
    document.querySelector(`.tab[data-tab="${tabName}"]`)?.classList.add("active");
    $("viewPunches").style.display = (tabName==="punches") ? "" : "none";
    $("viewDaily").style.display   = (tabName==="daily") ? "" : "none";
    $("viewSummary").style.display = (tabName==="summary") ? "" : "none";
    $("viewEmp").style.display     = (tabName==="emp") ? "" : "none";
    if(tabName==="emp") renderEmployeeReport();
  }
  document.querySelectorAll(".tab").forEach(t=> t.addEventListener("click", ()=> setActiveTab(t.getAttribute("data-tab"))) );

  // ===== Printing / PDF =====
  function nowString(){ const d=new Date(); return `${formatDate(d)} ${formatTime12(d)}`; }
  function currentRangeText(){ const m=$("monthFilter").value; if(m==="ALL") return "All months"; const [yyyy,mm]=m.split("-"); return `${mm}/${yyyy}`; }
  function tabTitle(tab){ return tab==="daily"?"Daily Attendance":tab==="summary"?"Monthly Summary":tab==="emp"?"Employee Report":"Cleaned Punches"; }
  function currentEmpText(){ const v=$("employeeFilter").value; if(v==="ALL") return "—"; const id=parseInt(v,10); const nm=schedule[id]?.name||""; return nm?`${id} - ${nm}`:String(id); }

  function fillPrintHeader(printTab){
    $("phCompany").textContent = ($("hdrCompany").value||"Attendance Report").trim();
    $("phPrepared").textContent = ($("hdrPrepared").value||"—").trim();
    $("phApproved").textContent = ($("hdrApproved").value||"—").trim();
    $("phNotes").textContent = ($("hdrNotes").value||"—").trim();
    $("phReportTitle").textContent = tabTitle(printTab);
    $("phRange").textContent = currentRangeText();
    $("phGen").textContent = nowString();
    $("phEmp").textContent = (printTab==="emp") ? currentEmpText() : "—";
  }

  
  // ===== Daily print: choose columns to hide =====
  const DAILY_COLS = ["EmployeeID","Name","ShiftDate","DayName","Status","SchedStart","SchedEnd","ActualIn","ActualOut","WorkMinutes","LateMinutes","OTMinutes","Present","Absent"];
  const DAILY_COL_LABELS = {
    EmployeeID: "المعرف",
    Name: "الاسم",
    ShiftDate: "التاريخ",
    DayName: "اليوم",
    Status: "الحالة",
    SchedStart: "بداية الدوام",
    SchedEnd: "نهاية الدوام",
    ActualIn: "دخول فعلي",
    ActualOut: "انصراف فعلي",
    WorkMinutes: "دقائق العمل",
    LateMinutes: "دقائق التأخير",
    OTMinutes: "دقائق الإضافي",
    Present: "حضور",
    Absent: "غياب"
  };
  const dailyHideState = new Set(); // column names to hide

  function renderDailyColPicker(){
    const host = $("dailyColGrid");
    if(!host) return;
    host.innerHTML = "";
    for(const col of DAILY_COLS){
      const lab = document.createElement("label");
      lab.className = "colItem";
      const ck = document.createElement("input");
      ck.type = "checkbox";
      ck.checked = dailyHideState.has(col);
      ck.addEventListener("change", ()=>{
        if(ck.checked) dailyHideState.add(col); else dailyHideState.delete(col);
        try{ localStorage.setItem("att_daily_hide", JSON.stringify(Array.from(dailyHideState))); }catch(_){}
      });
      const wrap = document.createElement("span");
      wrap.className = "colText";
      const ar = document.createElement("span");
      ar.className = "colAr";
      ar.textContent = (DAILY_COL_LABELS[col] || col);
      const en = document.createElement("span");
      en.className = "colEn";
      en.textContent = col;
      wrap.appendChild(ar); wrap.appendChild(en);
      lab.appendChild(ck);
      lab.appendChild(wrap);
      host.appendChild(lab);
    }
  }

  function setDailyMinimal(){
    // Minimal daily: keep essentials: EmployeeID, Name, ShiftDate, DayName, ActualIn, ActualOut, Status
    dailyHideState.clear();
    const keep = new Set(["EmployeeID","Name","ShiftDate","DayName","ActualIn","ActualOut","Status"]);
    for(const c of DAILY_COLS){ if(!keep.has(c)) dailyHideState.add(c); }
    renderDailyColPicker();
    try{ localStorage.setItem("att_daily_hide", JSON.stringify(Array.from(dailyHideState))); }catch(_){}
  }

  function resetDailyHide(){
    dailyHideState.clear();
    renderDailyColPicker();
    try{ localStorage.setItem("att_daily_hide", JSON.stringify([])); }catch(_){}
  }

  // Load persisted hide columns
  try{
    const saved = JSON.parse(localStorage.getItem("att_daily_hide") || "[]");
    if(Array.isArray(saved)) saved.forEach(c=> dailyHideState.add(String(c)));
  }catch(_){}

  function setPageOrientation(mode){
    const st = document.getElementById("pageStyle");
    if(!st) return;
    if(mode==="landscape"){
      st.textContent = "@page { size: A4 landscape; margin: 10mm; }";
    } else {
      st.textContent = "@page { size: A4 portrait; margin: 12mm; }";
    }
  }

  $("btnPrint").addEventListener("click", ()=>{
    const want = $("printWhat").value;
    const printTab = (want==="current") ? state.activeTab : want;
    const orientation = $("printOrientation").value;
    const fit = $("printFit").value;

    if(printTab==="emp" && $("employeeFilter").value==="ALL"){
      alert("للطباعة: اختر موظف محدد أولاً (Employee Report).");
      return;
    }

    setPageOrientation(orientation);
    fillPrintHeader(printTab);

    document.body.classList.add("printing");
    document.body.setAttribute("data-print-tab", printTab);
    document.body.setAttribute("data-fit", fit);

    // Apply daily hide columns only when printing Daily or Employee Report (employee daily table)
    if(printTab==="daily" || printTab==="emp"){
      const tokens = Array.from(dailyHideState).join(" ");
      if(tokens) document.body.setAttribute("data-hidecols", tokens);
      else document.body.removeAttribute("data-hidecols");
    } else {
      document.body.removeAttribute("data-hidecols");
    }

    window.print();
  });

  window.addEventListener("afterprint", ()=>{
    document.body.classList.remove("printing");
    document.body.removeAttribute("data-print-tab");
    document.body.removeAttribute("data-fit");
  });

  // ===== Theme toggle =====
  function setTheme(t){
    if(t==="light"){ document.body.setAttribute("data-theme","light"); $("btnTheme").textContent="☀️"; }
    else { document.body.removeAttribute("data-theme"); $("btnTheme").textContent="🌙"; }
    try{ localStorage.setItem("att_theme", t); }catch(_){}
  }
  $("btnTheme").addEventListener("click", ()=>{
    const cur = document.body.getAttribute("data-theme")==="light" ? "light" : "dark";
    setTheme(cur==="light" ? "dark" : "light");
  });
  try{
    const saved = localStorage.getItem("att_theme");
    if(saved==="light") setTheme("light");
  }catch(_){}

  

  // Daily print column picker init
  renderDailyColPicker();
  $("btnDailyMinimal")?.addEventListener("click", setDailyMinimal);
  $("btnDailyReset")?.addEventListener("click", resetDailyHide);
// ===== State =====
  let state = {
    rawCount:0, cleanCount:0,
    cleanPunches:[],
    punches:[], daily:[], summary:[],
    punchesFiltered:[], dailyFiltered:[],
    months:[], maxId:1,
    graceMin:15, rollHour:6,
    activeTab:"punches"
  };

  function applyFiltersAndRender(){
    const monthVal = $("monthFilter").value;

    state.daily = buildDailyWithAbsences(state.cleanPunches, state.rollHour, state.graceMin, monthVal, state.maxId);
    state.summary = buildSummary(state.daily);

    if(monthVal==="ALL"){
      state.punchesFiltered = state.punches;
      state.dailyFiltered = state.daily;
    } else {
      const [yyyy, mm] = monthVal.split("-").map(x=>parseInt(x,10));
      state.punchesFiltered = state.punches.filter(r=>{
        const parts = r.ShiftDate.split("/");
        return (+parts[2]===yyyy && +parts[1]===mm);
      });
      state.dailyFiltered = state.daily;
    }

    renderTable($("tblPunches"), state.punchesFiltered);
    renderTable($("tblDaily"), addTotalsRowDaily(state.dailyFiltered));
    renderTable($("tblSummary"), addTotalsRowSummary(state.summary));

    $("bestEmp").textContent = computeBest(state.summary);
    renderEmployeeReport();
  }

  // ===== Buttons =====
  $("btnProcess").addEventListener("click", ()=>{
    try{
      const rawText = $("raw").value || "";
      state.graceMin = Math.max(0, parseInt($("grace").value||"15",10));
      const dedupSeconds = Math.max(0, parseInt($("dedup").value||"60",10));
      state.rollHour = Math.max(0, Math.min(23, parseInt($("rollHour").value||"6",10)));

      const raw = parseRaw(rawText);
      state.rawCount = raw.length;
      if(raw.length===0){
        setStatus("لم يتم التعرف على أي صف. تأكد من dd/mm/yyyy وأن البيانات 5 أو 6 أعمدة.", true);
        $("kpi").style.display="none";
        renderTable($("tblPunches"), []);
        renderTable($("tblDaily"), []);
        renderTable($("tblSummary"), []);
        renderTable($("tblEmpDaily"), []);
        renderTable($("tblEmpPunches"), []);
        return;
      }

      const clean = dedupPunches(raw, dedupSeconds);
      state.cleanCount = clean.length;
      state.cleanPunches = clean;

      state.maxId = Math.max(...clean.map(p=>p.empId));
      if(!Number.isFinite(state.maxId) || state.maxId<1) state.maxId = 1;

      const nameMap = buildEmployeeNameMap(clean);
      ensureScheduleRows(nameMap, state.maxId);
      renderScheduleTable(state.maxId);

      state.punches = buildPunchRows(clean, state.rollHour);

      const tmpDailyAll = buildDailyWithAbsences(clean, state.rollHour, state.graceMin, "ALL", state.maxId);
      state.months = computeMonthsFromDaily(tmpDailyAll);

      populateMonthFilter(state.months);
      populateEmployeeFilter(state.maxId);

      applyFiltersAndRender();

      $("kpi").style.display="";
      $("kRaw").textContent=state.rawCount;
      $("kClean").textContent=state.cleanCount;
      $("kDaily").textContent=state.dailyFiltered.length;
      $("kEmp").textContent=state.maxId;

      setStatus("تمت المعالجة بنجاح ✅");
    }catch(e){
      console.error(e);
      setStatus("حدث خطأ أثناء المعالجة: " + (e?.message || e), true);
    }
  });

  $("btnClear").addEventListener("click", ()=>{
    $("raw").value="";
    setStatus("");
    $("kpi").style.display="none";
    state = {rawCount:0, cleanCount:0, cleanPunches:[], punches:[], daily:[], summary:[],
      punchesFiltered:[], dailyFiltered:[], months:[], maxId:1, graceMin:15, rollHour:6, activeTab:"punches"};
    renderTable($("tblPunches"), []);
    renderTable($("tblDaily"), []);
    renderTable($("tblSummary"), []);
    renderTable($("tblEmpDaily"), []);
    renderTable($("tblEmpPunches"), []);
    $("bestEmp").textContent="—";
    $("monthFilter").innerHTML='<option value="ALL">كل البيانات</option>';
    $("employeeFilter").innerHTML='<option value="ALL">كل الموظفين</option>';
    $("tblSchedule").innerHTML="";
    setActiveTab("punches");
  });

  $("monthFilter").addEventListener("change", ()=>{ if(state.cleanPunches.length) applyFiltersAndRender(); });
  $("employeeFilter").addEventListener("change", ()=>{ renderEmployeeReport(); });

  $("btnApplySchedule").addEventListener("click", ()=>{
    if(!state.cleanPunches.length){ alert("قم بعمل Process أولاً."); return; }
    if(!readScheduleFromUI(state.maxId)) return;
    applyFiltersAndRender();
    setStatus("تم تطبيق جدول الدوام وإعادة الحساب ✅");
  });

  $("btnExportSchedule").addEventListener("click", ()=>{
    if(state.maxId<1){ alert("قم بعمل Process أولاً."); return; }
    const rows=[];
    for(let id=1; id<=state.maxId; id++){
      const s=schedule[id];
      rows.push({
        EmployeeID:id,
        Name:s.name||"",
        IsManager:s.isManager?"YES":"NO",
        WeekdayStart:s.weekdayStart, WeekdayEnd:s.weekdayEnd,
        WeekendStart:s.weekendStart, WeekendEnd:s.weekendEnd,
        Sun:s.workDays[0]?1:0, Mon:s.workDays[1]?1:0, Tue:s.workDays[2]?1:0, Wed:s.workDays[3]?1:0,
        Thu:s.workDays[4]?1:0, Fri:s.workDays[5]?1:0, Sat:s.workDays[6]?1:0
      });
    }
    downloadTSV("schedule.tsv", rows);
  });

  $("btnImportSchedule").addEventListener("click", ()=>{
    if(!state.cleanPunches.length){ alert("قم بعمل Process أولاً."); return; }
    $("importScheduleFile").click();
  });

  $("importScheduleFile").addEventListener("change", async (ev)=>{
    const file = ev.target.files?.[0];
    if(!file) return;
    const text = await file.text();
    const lines = text.split(/\r?\n/).map(l=>l.trim()).filter(Boolean);
    if(lines.length<2){ alert("ملف الجدول فارغ."); return; }

    const header = lines[0].split("\t").map(x=>x.trim());
    const idx = (name)=> header.indexOf(name);
    const iId = idx("EmployeeID");
    if(iId<0){ alert("صيغة غير صحيحة: لا يوجد EmployeeID"); return; }

    const get = (cols, k)=>{ const j=idx(k); return j>=0 ? (cols[j]||"").trim() : ""; };
    let updated=0;

    for(let i=1;i<lines.length;i++){
      const cols = lines[i].split("\t");
      const empId = parseInt((cols[iId]||"").trim(),10);
      if(!Number.isFinite(empId) || empId<1 || empId>state.maxId) continue;

      const s = schedule[empId] || defaultScheduleFor(empId);

      const name = get(cols,"Name");
      const mgr  = get(cols,"IsManager");
      const ws   = get(cols,"WeekdayStart");
      const we   = get(cols,"WeekdayEnd");
      const wEs  = get(cols,"WeekendStart");
      const wEe  = get(cols,"WeekendEnd");
      const days = [get(cols,"Sun"),get(cols,"Mon"),get(cols,"Tue"),get(cols,"Wed"),get(cols,"Thu"),get(cols,"Fri"),get(cols,"Sat")]
        .map(x => String(x).trim()==="1" || String(x).trim().toUpperCase()==="TRUE" || String(x).trim().toUpperCase()==="YES");

      if(name) s.name=name;
      if(mgr) s.isManager = (mgr.toUpperCase()==="YES" || mgr==="1" || mgr.toUpperCase()==="TRUE");
      if(hhmmOk(ws)) s.weekdayStart=ws;
      if(hhmmOk(we)) s.weekdayEnd=we;
      if(hhmmOk(wEs)) s.weekendStart=wEs;
      if(hhmmOk(wEe)) s.weekendEnd=wEe;
      if(days.length===7) s.workDays=days;

      schedule[empId]=s;
      updated++;
    }

    renderScheduleTable(state.maxId);
    applyFiltersAndRender();
    alert(`تم استيراد/تحديث ${updated} موظف.`);
    ev.target.value="";
  });

  // exports
  $("btnExportPunches").addEventListener("click", ()=>{
    if(!state.punchesFiltered.length){ alert("لا توجد بيانات."); return; }
    const m=$("monthFilter").value;
    downloadTSV(`punches_${m==="ALL"?"all":m}.tsv`, state.punchesFiltered);
  });
  $("btnExportDaily").addEventListener("click", ()=>{
    if(!state.dailyFiltered.length){ alert("لا توجد بيانات."); return; }
    const m=$("monthFilter").value;
    downloadTSV(`daily_${m==="ALL"?"all":m}.tsv`, addTotalsRowDaily(state.dailyFiltered));
  });
  $("btnExportSummary").addEventListener("click", ()=>{
    if(!state.summary.length){ alert("لا توجد بيانات."); return; }
    const m=$("monthFilter").value;
    downloadTSV(`summary_${m==="ALL"?"all":m}.tsv`, addTotalsRowSummary(state.summary));
  });

  
  $("btnXlsxSummary").addEventListener("click", ()=> exportDatasetXLSX("summary"));
  $("btnXlsxDaily").addEventListener("click", ()=> exportDatasetXLSX("daily"));
  $("btnXlsxPunches").addEventListener("click", ()=> exportDatasetXLSX("punches"));
  $("btnXlsxEmployee").addEventListener("click", ()=> exportDatasetXLSX("employee"));
// Init
  renderTable($("tblPunches"), []);
  renderTable($("tblDaily"), []);
  renderTable($("tblSummary"), []);
  renderTable($("tblEmpDaily"), []);
  renderTable($("tblEmpPunches"), []);
  setActiveTab("punches");
})();
