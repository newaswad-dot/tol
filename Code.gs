/*****************************
 * إعدادات عامة + قراءة Settings
 *****************************/
function getConfig_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Settings");
  if (!sh) throw new Error("❌ لم يتم العثور على ورقة Settings. أنشئ ورقة باسم Settings وضع القيم في الصف 2.");

  // الصف2:
  // A=AGENT_SHEET_ID, B=AGENT_SHEET_NAME, C=ADMIN_SHEET_ID, D=ADMIN_SHEET_NAME
  // E=DATA1_ID, F=DATA1_NAME, G=DATA2_ID, H=DATA2_NAME  (اختياري)
  const row = sh.getRange(2, 1, 1, 8).getValues()[0];
  const cfg = {
    AGENT_SHEET_ID:   String(row[0] || "").trim(),
    AGENT_SHEET_NAME: String(row[1] || "").trim() || "SHEET",
    ADMIN_SHEET_ID:   String(row[2] || "").trim(),
    ADMIN_SHEET_NAME: String(row[3] || "").trim() || "Sheet1",

    DATA1_ID:         String(row[4] || "").trim(),
    DATA1_NAME:       String(row[5] || "").trim() || "معلومات السلطان",
    DATA2_ID:         String(row[6] || "").trim(),
    DATA2_NAME:       String(row[7] || "").trim() || "معلومات الفرعيين",
  };
  const missing = [];
  if (!cfg.AGENT_SHEET_ID)   missing.push("AGENT_SHEET_ID");
  if (!cfg.AGENT_SHEET_NAME) missing.push("AGENT_SHEET_NAME");
  if (!cfg.ADMIN_SHEET_ID)   missing.push("ADMIN_SHEET_ID");
  if (!cfg.ADMIN_SHEET_NAME) missing.push("ADMIN_SHEET_NAME");
  if (missing.length) throw new Error("⚠️ إعدادات ناقصة في Settings: " + missing.join(", "));
  return cfg;
}

function getConfigStatus() {
  try { return { ok:true, config:getConfig_() }; }
  catch(e){ return { ok:false, message:e.message }; }
}

/*****************************
 * كاش + مفاتيح
 *****************************/
const CACHE_TTL_SEC       = 21600; // 6 ساعات
const KEY_AGENT_INDEX     = "agentIndex_v8";   // { [id]: { rows:[..], names:[..], salaries:[..], sum:number } } - تم التحديث
const KEY_ADMIN_IDSET     = "adminIdSet_v7";   // { [id]:1 }
const KEY_ADMIN_ROW_MAP   = "adminRowMap_v7";  // { [id]: [rowIndex,...] }
const KEY_COLORED_AGENT   = "coloredAgentIds_v7";
const KEY_COLORED_ADMIN   = "coloredAdminIds_v7";
const KEY_CORR_MAP        = "salaryCorrMap_v1"; // { "30":29, "88":82, ... }
// كاش معلومات الأشخاص:
const KEY_INFO_ID2GROUP   = "info_id2group_v1"; // { id: groupKey }
const KEY_INFO_GROUPS     = "info_groups_v1";   // { groupKey: {...} }

/********* أدوات chunk للكاش *********/
function cachePutChunked_(keyPrefix, obj, cache) {
  const txt = JSON.stringify(obj);
  const MAX = 90000;
  const n   = Math.ceil(txt.length / MAX);
  const bag = {};
  for (let i = 0; i < n; i++) bag[`${keyPrefix}_chunk_${i}`] = txt.substring(i*MAX, (i+1)*MAX);
  bag[`${keyPrefix}_chunk_count`] = String(n);
  cache.putAll(bag, CACHE_TTL_SEC);
}
function cacheGetChunked_(keyPrefix, cache) {
  const c = cache.get(`${keyPrefix}_chunk_count`);
  if (!c) return null;
  const n = parseInt(c,10);
  const keys = Array.from({length:n},(_,i)=>`${keyPrefix}_chunk_${i}`);
  const got  = cache.getAll(keys);
  let out = "";
  for (let i=0;i<n;i++){
    const part = got[`${keyPrefix}_chunk_${i}`];
    if (!part) return null;
    out += part;
  }
  try { return JSON.parse(out); } catch(_) { return null; }
}

/*****************************
 * بناء الفهارس (وكيل/إدارة)
 *****************************/
function buildAgentIndex_(colA, colB, colC) {
  const index = Object.create(null);
  const n = Math.max(colA.length, colB.length, colC.length);
  for (let i=0;i<n;i++){
    const id  = String(colA[i] || '').trim();
    if (!id) continue;
    const name = String(colB[i] || '').trim(); // قراءة الاسم من العمود B
    const sal = parseFloat(colC[i] || 0);
    if (!index[id]) index[id] = { rows:[], names:[], salaries:[], sum:0 };
    index[id].rows.push(i+1); // 1-based
    index[id].names.push(name); // تخزين الاسم
    const s = isNaN(sal) ? 0 : sal;
    index[id].salaries.push(s);
    index[id].sum += s;
  }
  return index;
}

function buildColoredIdSet_(ssId, sheetName) {
  const ss = SpreadsheetApp.openById(ssId);
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return {};
  const lastRow = sh.getLastRow();
  if (lastRow < 1) return {};
  const colA = sh.getRange(1,1,lastRow,1).getDisplayValues().flat();
  const bgs  = sh.getRange(1,1,lastRow,1).getBackgrounds().flat();
  const set = Object.create(null);
  for (let i=0;i<colA.length;i++){
    const id = String(colA[i]||'').trim();
    if (!id) continue;
    const c = String(bgs[i]||'').toLowerCase();
    if (c && c !== '#ffffff' && c !== 'white' && c !== 'transparent') set[id] = 1;
  }
  return set;
}

/*****************************
 * ورقة "تصحيح الراتب" (اختيارية)
 * أعمدة: A=الراتب الأصلي، B=الراتب المعروض
 *****************************/
function buildSalaryCorrectionsMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("تصحيح الراتب");
  const map = Object.create(null);
  if (!sh) return map;
  const lr = sh.getLastRow();
  if (lr < 1) return map;
  const vals = sh.getRange(1,1,lr,2).getDisplayValues();
  for (let i=0;i<vals.length;i++){
    const from = Number(vals[i][0]);
    const to   = Number(vals[i][1]);
    if (!isNaN(from) && !isNaN(to)) map[String(from)] = to;
  }
  return map;
}
function applySalaryCorrection_(val, corrMap) {
  const key = String(Number(val));
  if (corrMap && Object.prototype.hasOwnProperty.call(corrMap, key)) {
    return Number(corrMap[key]);
  }
  return Number(val||0);
}

/*****************************
 * أدوات “معلومات الأشخاص”
 *****************************/
function openSheetFlex_(idMaybe, nameMaybe) {
  if (idMaybe) {
    const ss = SpreadsheetApp.openById(idMaybe);
    if (nameMaybe) {
      const sh = ss.getSheetByName(nameMaybe);
      if (sh) return sh;
    }
    return ss.getSheets()[0] || null;
  } else {
    const cur = SpreadsheetApp.getActiveSpreadsheet();
    if (nameMaybe) return cur.getSheetByName(nameMaybe);
    return null;
  }
}
function readSheetAsObjectsWithSource_(sh, sourceKey) {
  if (!sh) return [];
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return [];
  const headers = sh.getRange(1,1,1,lc).getDisplayValues()[0].map(h=>String(h||'').trim());
  const vals = sh.getRange(2,1,lr-1,lc).getDisplayValues();
  return vals.map(row=>{
    const obj = {};
    for (let i=0;i<headers.length;i++){
      const key = headers[i] || ('COL_'+(i+1));
      obj[key] = row[i];
    }
    obj.__source = sourceKey; // data1 أو data2
    return obj;
  });
}
// نسخة بدون مصدر (تُستخدم في applyAdvancedAction)
function readSheetAsObjects_(sh) {
  if (!sh) return [];
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return [];
  const headers = sh.getRange(1,1,1,lc).getDisplayValues()[0].map(h=>String(h||'').trim());
  const vals = sh.getRange(2,1,lr-1,lc).getDisplayValues();
  return vals.map(row=>{
    const obj = {};
    for (let i=0;i<headers.length;i++){
      const key = headers[i] || ('COL_'+(i+1));
      obj[key] = row[i];
    }
    return obj;
  });
}

function normalizeName_(s){ return String(s||'').trim().replace(/\s+/g,' ').toLowerCase(); }
function extractNameFromRow_(row){
  if (!row) return '';
  const keys = ['full_name','الاسم','الاسم الثلاثي','name'];
  for (const k in row){
    if (keys.indexOf(String(k).toLowerCase()) !== -1){
      const v = String(row[k]||'').trim();
      if (v) return v;
    }
  }
  return '';
}
function pickField_(row, keyAliases, defVal){
  for (const k in row) {
    const kl = k.toLowerCase();
    if (keyAliases.indexOf(kl) !== -1) {
      const v = String(row[k]||'').trim();
      if (v) return v;
    }
  }
  return defVal;
}
function extractIdsFromRow_(obj) {
  const out = {};
  for (const k in obj) {
    const kl = String(k).toLowerCase();
    if (kl === 'id') {
      String(obj[k]||'').split(/[,\s]+/).forEach(function(x){
        const v = String(x||'').trim();
        if (v) out[v] = 1;
      });
    }
    if (kl === 'raw_payload_json') {
      try{
        const j = JSON.parse(obj[k]||'{}');
        const arr = j && j.user_ids;
        if (Array.isArray(arr)) arr.forEach(function(v){ const s=String(v||'').trim(); if(s) out[s]=1; });
      }catch(_){}
    }
  }
  return Object.keys(out);
}
function buildInfoGroups_(rows1, rows2){
  const groups = Object.create(null);
  const id2group = Object.create(null);
  let anonCounter = 0;

  function mergeRow(r){
    const src = r.__source || '';
    const name = extractNameFromRow_(r);
    const gk0  = normalizeName_(name);
    const gk   = gk0 || ('__anon__'+(++anonCounter));

    if (!groups[gk]){
      const phone   = pickField_(r, ['phone','الهاتف','رقم الهاتف','mobile'], '');
      const address = pickField_(r, ['address','العنوان','المحافظة','المدينة'], '');
      const agency  = pickField_(r, ['agency_name','الوكالة','الشركة','الفرع'], '');
      const noteLbl = pickField_(r, ['extra_field_label'],'');
      const noteVal = pickField_(r, ['extra_field_value'],'');
      const note    = (noteLbl && noteVal) ? (noteLbl+' : '+noteVal) : (noteVal || '');

      groups[gk] = {
        name: name || '',
        phone, address, agency, note,
        ids: [],
        sources: { data1:false, data2:false }
      };
    }

    const ids = extractIdsFromRow_(r);
    for (let i=0;i<ids.length;i++){
      const id = String(ids[i]).trim();
      if (!id) continue;
      if (!groups[gk].ids.some(x=>x.id===id)){
        groups[gk].ids.push({ id:id, source: src });
      }
      if (!id2group[id]) id2group[id] = gk;
    }
    if (src === 'data1') groups[gk].sources.data1 = true;
    if (src === 'data2') groups[gk].sources.data2 = true;

    function enrich(fieldName, val){
      if (!groups[gk][fieldName] && val) groups[gk][fieldName] = val;
    }
    enrich('phone',   pickField_(r, ['phone','الهاتف','رقم الهاتف','mobile'], ''));
    enrich('address', pickField_(r, ['address','العنوان','المحافظة','المدينة'], ''));
    enrich('agency',  pickField_(r, ['agency_name','الوكالة','الشركة','الفرع'], ''));
    if (!groups[gk].note){
      const nl = pickField_(r, ['extra_field_label'],'');
      const nv = pickField_(r, ['extra_field_value'],'');
      const nt = (nl && nv) ? (nl+' : '+nv) : (nv || '');
      if (nt) groups[gk].note = nt;
    }
  }

  (rows1||[]).forEach(mergeRow);
  (rows2||[]).forEach(mergeRow);

  return { groups, id2group };
}

/*****************************
 * تحميل البيانات إلى الكاش
 *****************************/
function loadDataIntoCache() {
  try {
    const cache = CacheService.getScriptCache();
    const cfg = getConfig_();

    // الوكيل
    const agSS = SpreadsheetApp.openById(cfg.AGENT_SHEET_ID);
    const agSh = agSS.getSheetByName(cfg.AGENT_SHEET_NAME);
    if (!agSh) throw new Error('لم يتم العثور على ورقة الوكيل "'+cfg.AGENT_SHEET_NAME+'".');
    const agLastRow = agSh.getLastRow();
    let agentIndex = {};
    if (agLastRow > 0) {
      const colA = agSh.getRange(1,1,agLastRow,1).getValues().flat(); // IDs
      const colB = agSh.getRange(1,2,agLastRow,1).getValues().flat(); // الأسماء - العمود B
      const colC = agSh.getRange(1,3,agLastRow,1).getValues().flat(); // الرواتب
      agentIndex = buildAgentIndex_(colA, colB, colC); // تمرير الأعمدة الثلاثة
    }

    // الإدارة
    const adSS = SpreadsheetApp.openById(cfg.ADMIN_SHEET_ID);
    const adSh = adSS.getSheetByName(cfg.ADMIN_SHEET_NAME);
    if (!adSh) throw new Error('لم يتم العثور على ورقة الإدارة "'+cfg.ADMIN_SHEET_NAME+'".');
    const adLastRow = adSh.getLastRow();
    let adminIdSet = {}, adminRowMap = {};
    if (adLastRow > 0) {
      const colA = adSh.getRange(1,1,adLastRow,1).getValues().flat(); // IDs
      for (let i=0; i<colA.length; i++) {
        const id = String(colA[i]||'').trim();
        if (!id) continue;
        adminIdSet[id] = 1;
        if (!adminRowMap[id]) adminRowMap[id] = [];
        adminRowMap[id].push(i+1); // 1-based
      }
    }

    // خرائط الملوّن
    const coloredAgent = buildColoredIdSet_(cfg.AGENT_SHEET_ID, cfg.AGENT_SHEET_NAME);
    const coloredAdmin = buildColoredIdSet_(cfg.ADMIN_SHEET_ID, cfg.ADMIN_SHEET_NAME);

    // خريطة تصحيح الراتب
    const corrMap = buildSalaryCorrectionsMap_();

    // شيتات معلومات الأشخاص
    const sh1 = openSheetFlex_(cfg.DATA1_ID, cfg.DATA1_NAME); // معلومات السلطان
    const sh2 = openSheetFlex_(cfg.DATA2_ID, cfg.DATA2_NAME); // الفرعيين
    const rows1 = readSheetAsObjectsWithSource_(sh1, 'data1');
    const rows2 = readSheetAsObjectsWithSource_(sh2, 'data2');
    const infoPacked = buildInfoGroups_(rows1, rows2); // { groups, id2group }

    // اكتب في الكاش (chunked)
    cachePutChunked_(KEY_AGENT_INDEX,   agentIndex, cache);
    cachePutChunked_(KEY_ADMIN_IDSET,   adminIdSet, cache);
    cachePutChunked_(KEY_ADMIN_ROW_MAP, adminRowMap,cache);
    cachePutChunked_(KEY_COLORED_AGENT, coloredAgent, cache);
    cachePutChunked_(KEY_COLORED_ADMIN, coloredAdmin, cache);
    cachePutChunked_(KEY_CORR_MAP,      corrMap,     cache);
    cachePutChunked_(KEY_INFO_ID2GROUP, infoPacked.id2group, cache);
    cachePutChunked_(KEY_INFO_GROUPS,   infoPacked.groups,   cache);

    // إحصاء بسيط
    let agentRows = 0;
    for (const id in agentIndex) agentRows += (agentIndex[id].rows ? agentIndex[id].rows.length : 0);
    const agentUnique = Object.keys(agentIndex).length;

    let adminRows = 0;
    for (const id in adminRowMap) adminRows += (adminRowMap[id] ? adminRowMap[id].length : 0);

    return {
      success:true,
      message:'تم التحميل ✓ — الوكيل: '+agentRows+' صف / '+agentUnique+' ID فريد — الإدارة: '+adminRows+' صف.'
    };
  } catch (e) {
    return { success:false, message:'خطأ: ' + e.message };
  }
}

/*****************************
 * سنابشوت محلي سريع للواجهة
 *****************************/
function getSearchSnapshotLight() {
  try {
    const cache = CacheService.getScriptCache();
    const agentIndex   = cacheGetChunked_(KEY_AGENT_INDEX,   cache) || {};
    const adminIdSet   = cacheGetChunked_(KEY_ADMIN_IDSET,   cache) || {};
    const coloredAgent = cacheGetChunked_(KEY_COLORED_AGENT, cache) || {};
    const coloredAdmin = cacheGetChunked_(KEY_COLORED_ADMIN, cache) || {};

    const map = {};
    let agentRows = 0;
    for (const id in agentIndex) {
      const node = agentIndex[id] || {};
      const rowsLen = (node.rows && node.rows.length) ? node.rows.length : 0;
      agentRows += rowsLen;
      map[id] = {
        sum: Number(node.sum||0),
        salaries: (node.salaries||[]).map(Number),
        names: (node.names||[]).slice(), // نقل الأسماء
        rowsCount: rowsLen,
        inAdmin: !!adminIdSet[id],
        aCol: !!coloredAgent[id],
        dCol: !!coloredAdmin[id]
      };
    }
    return { ok:true, map:map, stats:{ agentRows:agentRows, agentUnique:Object.keys(agentIndex).length } };
  } catch(e){
    return { ok:false, message:e.message };
  }
}

/*****************************
 * بحث سريع + ملخص
 *****************************/
function searchId(id, discountPercentage) {
  try {
    if (!id) return { status:'error', message:'الرجاء إدخال ID للبحث.' };
    id = String(id).trim();

    const cache = CacheService.getScriptCache();
    const agentIndex   = cacheGetChunked_(KEY_AGENT_INDEX,   cache);
    const adminIdSet   = cacheGetChunked_(KEY_ADMIN_IDSET,   cache);
    const coloredAgent = cacheGetChunked_(KEY_COLORED_AGENT, cache);
    const coloredAdmin = cacheGetChunked_(KEY_COLORED_ADMIN, cache);

    if (!agentIndex || !adminIdSet || !coloredAgent || !coloredAdmin) {
      return { status:'error', message:'البيانات غير محمّلة. اضغط "تحميل البيانات".' };
    }

    const inAgent = !!agentIndex[id];
    const inAdmin = !!adminIdSet[id];

    // ← مهم: نعرّف total من البداية ونستخدمه لاحقًا أينما كان الفرع
    let status   = 'غير موجود';
    let salaries = [];
    let names    = [];
    let total    = 0;

    if (inAgent) {
      const node = agentIndex[id];
      salaries = (node.salaries || []).slice();
      names    = (node.names || []).slice();
      total    = Number(node.sum || 0);
      status   = inAdmin
        ? ((node.rows.length > 1) ? 'سحب وكالة - راتبين' : 'سحب وكالة')
        : ((node.rows.length > 1) ? 'راتبين' : 'وكالة');
    } else if (inAdmin) {
      status = 'ادارة';
      total  = 0; // ← حتى لو إدارة فقط يبقى total معرّف
    } else {
      // غير موجود في الاثنين
      return {
        status:'غير موجود',
        totalSalary:'0.00',
        salaries:[],
        names:[],
        name:'',
        discountAmount:'0.00',
        salaryAfterDiscount:'0.00',
        id:id,
        isDuplicate:false
      };
    }

    // مكرر؟
    let isDuplicate = false, duplicateLabel = null;
    const aCol = !!coloredAgent[id];
    const dCol = !!coloredAdmin[id];
    if (aCol && dCol)      { isDuplicate = true; duplicateLabel = 'مكرر'; }
    else if (aCol)         { isDuplicate = true; duplicateLabel = 'مكرر وكالة فقط'; }
    else if (dCol)         { isDuplicate = true; duplicateLabel = 'مكرر ادارة فقط'; }

    // اسم مختصر للواجهة
    const primaryName = (names && names.length) ? String(names[0] || '').trim() : '';

    // الخصم
    const p    = Math.max(0, Math.min(100, Number(discountPercentage) || 0));
    const disc = total * (p / 100);
    const aft  = total - disc;

    return {
      status: status,
      totalSalary: total.toFixed(2),
      salaries: salaries,
      names: names,
      name: primaryName,
      discountAmount: disc.toFixed(2),
      salaryAfterDiscount: aft.toFixed(2),
      id: id,
      isDuplicate: isDuplicate,
      duplicateLabel: duplicateLabel
    };
  } catch (e) {
    return { status:'error', message: e.toString() };
  }
}

function getLiveStatsForFooter(discountPercentage) {
  try {
    const cache = CacheService.getScriptCache();
    const agentIndex   = cacheGetChunked_(KEY_AGENT_INDEX,   cache) || {};
    const coloredAgent = cacheGetChunked_(KEY_COLORED_AGENT, cache) || {};

    let totalRowsWithIds = 0;
    let coloredRows = 0;
    let totalSalary = 0;
    let multiRows = 0;

    for (const id in agentIndex) {
      const node = agentIndex[id] || {};
      const rowsCount = (node.rows && node.rows.length) ? node.rows.length : 0;
      totalRowsWithIds += rowsCount;
      if (coloredAgent[id]) coloredRows += rowsCount;
      totalSalary += Number(node.sum || 0);
      if (rowsCount > 1) multiRows++;
    }
    const uncoloredRows = Math.max(0, totalRowsWithIds - coloredRows);

    const p = Math.max(0, Math.min(100, Number(discountPercentage)||0));
    const totalDiscount = totalSalary * (p/100);
    const afterDiscount = totalSalary - totalDiscount;

    return {
      ok: true,
      agentIdCount: Object.keys(agentIndex).length,
      coloredRows: coloredRows,
      uncoloredRows: uncoloredRows,
      multiRows: multiRows,
      totalSalary: Number(totalSalary.toFixed(2)),
      discountPercent: p,
      totalDiscount: Number(totalDiscount.toFixed(2)),
      afterDiscount: Number(afterDiscount.toFixed(2))
    };
  } catch (e) {
    return { ok:false, message: e.message || String(e) };
  }
}

/*****************************
 * بطاقة الشخص + تصحيح راتب (يحترم السويتش)
 *****************************/
function buildPersonCardFromGroup_(group, agentIndex, corrMap) {
  const props = PropertiesService.getScriptProperties();
  const useCorr = (props.getProperty('USE_SAL_CORR') === '1');

  const name    = String(group.name||'').trim() || '—';
  const phone   = String(group.phone||'').trim() || '—';
  const address = String(group.address||'').trim() || '—';
  const agency  = String(group.agency||'').trim() || '—';
  const note    = String(group.note||'').trim();

  const idLines = [];
  let total = 0;
  const ids = Array.isArray(group.ids) ? group.ids : [];
  for (let i=0;i<ids.length;i++){
    const uid = ids[i].id;
    const node = agentIndex && agentIndex[uid];
    const sumOrig = node ? Number(node.sum||0) : 0;
    const sumShown = useCorr ? applySalaryCorrection_(sumOrig, corrMap) : sumOrig;
    total += sumShown;
    idLines.push({ id: uid, amount: sumShown });
  }

  return {
    ok:true,
    name: name, phone: phone, address: address, agency: agency, note: note,
    ids: idLines,
    total: total,
    sources: {
      data1: !!(group.sources && group.sources.data1),
      data2: !!(group.sources && group.sources.data2)
    }
  };
}

function getPersonCardById(id) {
  try{
    id = String(id||'').trim();
    if (!id) return { ok:false, message:'أدخل ID' };

    const cache = CacheService.getScriptCache();
    const id2group   = cacheGetChunked_(KEY_INFO_ID2GROUP, cache) || {};
    const groups     = cacheGetChunked_(KEY_INFO_GROUPS,   cache) || {};
    const agentIndex = cacheGetChunked_(KEY_AGENT_INDEX,   cache) || {};
    const corrMap    = cacheGetChunked_(KEY_CORR_MAP,      cache) || {};
    const coloredAgent = cacheGetChunked_(KEY_COLORED_AGENT, cache) || {};
    const coloredAdmin = cacheGetChunked_(KEY_COLORED_ADMIN, cache) || {};

    if (!id2group || !groups || !agentIndex) {
      return { ok:false, message:'⚠️ البيانات غير محمّلة. اضغط "تحميل البيانات".' };
    }

    const gk = id2group[id];
    if (!gk || !groups[gk]) {
      return { ok:false, message:'لم يتم العثور على بيانات لهذا ID في شيتات المعلومات.' };
    }

    const card = buildPersonCardFromGroup_(groups[gk], agentIndex, corrMap);

    const duplicates = [];
    for (let i=0;i<card.ids.length;i++){
      const uid = card.ids[i].id;
      const aCol = !!coloredAgent[uid];
      const dCol = !!coloredAdmin[uid];
      if (aCol || dCol) {
        const label = (aCol && dCol) ? 'مكرر (وكالة + إدارة)' : (aCol ? 'مكرر وكالة' : 'مكرر إدارة');
        duplicates.push({ id: uid, label: label });
      }
    }

    return Object.assign({}, card, { duplicates: duplicates });
  }catch(e){
    return { ok:false, message: 'خطأ: ' + (e.message||String(e)) };
  }
}

/*****************************
 * أدوات مساعدة للـ applyAdvancedAction
 *****************************/
function findProfileRowById_(rows, id) {
  id = String(id||'').trim();
  if (!id || !Array.isArray(rows)) return null;
  for (let i=0;i<rows.length;i++){
    const r = rows[i];
    const ids = extractIdsFromRow_(r) || [];
    if (ids.indexOf(id) !== -1) {
      return { rowIndex: i+2, allIds: ids }; // 2 = بعد صف العناوين
    }
  }
  return null;
}

/***** === FAST RANGE COLORING (bulk contiguous runs) === *****/
function colorRowsFast_(sh, rows, bg) {
  try {
    if (!sh || !rows || !rows.length) return;
    var lastCol = sh.getLastColumn();
    var color = bg || '#ddd6fe';

    // رتّب وجمّع الصفوف المتتالية لتقليل عدد اللمسات
    rows = rows.slice().sort(function(a,b){return a-b;});
    var runs = [];
    var start = rows[0], prev = rows[0];
    for (var i=1; i<rows.length; i++){
      var r = rows[i];
      if (r === prev + 1) { prev = r; continue; }
      runs.push([start, prev]); start = prev = r;
    }
    runs.push([start, prev]);

    // لوّن كل مقطع بلمسة واحدة باستخدام setBackgrounds
    for (var k=0; k<runs.length; k++){
      var s = runs[k][0], e = runs[k][1];
      var h = e - s + 1;
      var rng = sh.getRange(s, 1, h, lastCol);
      var block = Array.from({length:h}, function(){ return Array(lastCol).fill(color); });
      rng.setBackgrounds(block);
    }
  } catch(e) {}
}

/*****************************
 * تنفيذ ذكي + نسخ/تلوين (مع منع تكرار مُحكم)
 *****************************/
function applyAdvancedAction(id, targetSheet, adminColor, withdrawColor, targetMode, expandAllProfileIds) {
  try {
    id = String(id||'').trim();
    if(!id) return {success:false,message:"❌ أدخل ID"};

    targetMode = (targetMode||'both').toLowerCase(); 
    const doAdminOps = (targetMode === 'both');     // "الإدارة + الوكيل"
    expandAllProfileIds = (expandAllProfileIds !== false); // افتراضي: يوسّع

    const cache = CacheService.getScriptCache();
    const agentIndex  = cacheGetChunked_(KEY_AGENT_INDEX,   cache) || {};
    const adminRowMap = cacheGetChunked_(KEY_ADMIN_ROW_MAP, cache) || {};
    let coloredAgent  = cacheGetChunked_(KEY_COLORED_AGENT, cache) || {};
    let coloredAdmin  = cacheGetChunked_(KEY_COLORED_ADMIN, cache) || {};

    const cfg = getConfig_();
    const adSS = SpreadsheetApp.openById(cfg.ADMIN_SHEET_ID);
    const adSh = adSS.getSheetByName(cfg.ADMIN_SHEET_NAME);

    let tgSh = null, targetIdSet = Object.create(null);
    if (doAdminOps) {
      tgSh = adSS.getSheetByName(targetSheet || '');
      if(!tgSh) return {success:false,message:"⚠️ اختر ورقة الهدف أولاً"};

      // IDs الموجودة مسبقًا في ورقة الهدف (لمنع التكرار)
      const lr = tgSh.getLastRow();
      if (lr > 0) {
        const colA = tgSh.getRange(1,1,lr,1).getDisplayValues();
        for (var i=0;i<lr;i++){
          const cur = (colA[i][0]||'').trim();
          if (cur) targetIdSet[cur] = 1;
        }
      }
    }

    // توسيع IDs الخاصّة بالشخص (اختياري)
    let targetIds = [id];
    if (expandAllProfileIds) {
      const sh1 = openSheetFlex_(cfg.DATA1_ID, cfg.DATA1_NAME);
      const sh2 = openSheetFlex_(cfg.DATA2_ID, cfg.DATA2_NAME);
      const rows1 = readSheetAsObjects_(sh1);
      const rows2 = readSheetAsObjects_(sh2);
      let found = findProfileRowById_(rows1, id) || findProfileRowById_(rows2, id);
      if (found && Array.isArray(found.allIds) && found.allIds.length) {
        targetIds = Array.from(new Set(found.allIds.map(String)));
      }
    }

    let copied = 0, skipped = 0, totalColored = 0;
    let lastUsedColor = null;

    const agSS = SpreadsheetApp.openById(cfg.AGENT_SHEET_ID);
    const agSh = agSS.getSheetByName(cfg.AGENT_SHEET_NAME);

    // منع تكرار نفس الـID ضمن نفس ضغطة التنفيذ
    const recentCopied = Object.create(null);

    // انسخ "صف واحد فقط" من الإدارة إلى الهدف (منع تكرار ذكي بدون منع بسبب التلوين)
function copyOneAdminRowToTarget(adRows, colorHex){
  if (!doAdminOps || !tgSh || !Array.isArray(adRows) || !adRows.length) return;

  const adLastCol = adSh.getLastColumn();

  for (let i = 0; i < adRows.length; i++){
    const r = adRows[i];
    const vals = adSh.getRange(r, 1, 1, adLastCol).getValues()[0];
    const curIdFromRow = String(vals[0] || '').trim();

    if (alreadyCopied_(curIdFromRow, tgSh)) { skipped++; continue; }
    
    // ✋ منع التكرار الحقيقي فقط:
    // - موجود مسبقًا في ورقة الهدف
    // - أو تكرر ضمن نفس ضغطة التنفيذ
    if (!curIdFromRow) { skipped++; continue; }
    if (targetIdSet[curIdFromRow]) { skipped++; continue; }
    if (recentCopied && recentCopied[curIdFromRow]) { skipped++; continue; }

    // مكان اللصق
    const startAt = tgSh.getLastRow() + 1;

    // اللصق
    tgSh.appendRow(vals);

    // تلوين الصف الملصوق
    const lastColTarget = tgSh.getLastColumn() || adLastCol;
    tgSh.getRange(startAt, 1, 1, lastColTarget).setBackground(colorHex);

    // حدّث منع التكرار
    targetIdSet[curIdFromRow] = 1;          // صار موجود بالهدف
    if (typeof recentCopied === 'object') {
      recentCopied[curIdFromRow] = 1;       // لا تكرّره ضمن نفس الضغطة
    }

    copied++;

    // (اختياري) سجل العملية
    try { logCopyOperation(tgSh.getName(), startAt, 1, colorHex, curIdFromRow); } catch(_){}

    break; // صف واحد فقط لكل ID
  }
}

    // تنفيذ لكل ID
    targetIds.forEach(function(curId){
      const node    = agentIndex[curId];
      const inAgent = !!node;
      const agRows  = (node && node.rows) || [];
      const adRows  = adminRowMap[curId] || [];
      const inAdmin = adRows.length > 0;

      if (!inAgent && !inAdmin) return;

      // تحديد الحالة
      let status;
      if (inAgent && inAdmin) status = (agRows.length>1) ? 'سحب وكالة - راتبين' : 'سحب وكالة';
      else if (inAdmin)       status = 'ادارة';
      else                    status = (agRows.length>1) ? 'راتبين' : 'وكالة';

      // === إدارة فقط ===
      if (!status.includes('سحب وكالة') && status.includes('ادارة')) {
        if (doAdminOps && adRows.length){
          const usedColor = adminColor || '#fde68a';

          // لوّن الإدارة مرة واحدة فقط لكل ID (منع إعادة التلوين)
          if (!coloredAdmin[curId]) {
            colorRowsFast_(adSh, adRows, usedColor);
            coloredAdmin[curId] = 1;
            totalColored += adRows.length;
            lastUsedColor = usedColor;
          }

          // انسخ صف واحد فقط من الإدارة إلى الهدف (مع منع التكرار)
          copyOneAdminRowToTarget(adRows, usedColor);
        }
      }

      // === "سحب وكالة" أو "وكالة/راتبين" ===
      if (status.includes('سحب وكالة') || status.includes('وكالة')) {
        const usedColor = withdrawColor || '#ddd6fe';

        // لوّن الوكيل مرة واحدة فقط لكل ID
        if (agRows.length && !coloredAgent[curId]){
          colorRowsFast_(agSh, agRows, usedColor);
          coloredAgent[curId] = 1;
          totalColored += agRows.length;
          lastUsedColor = usedColor;
        }

        // إذا "الإدارة + الوكيل": لوّن الإدارة مرة واحدة + انسخ صف واحد
        if (doAdminOps && adRows.length){
          if (!coloredAdmin[curId]) {
            colorRowsFast_(adSh, adRows, usedColor);
            coloredAdmin[curId] = 1;
            totalColored += adRows.length;
            lastUsedColor = usedColor;
          }
          copyOneAdminRowToTarget(adRows, usedColor);
        }
      }
    });

    SpreadsheetApp.flush();
    cachePutChunked_(KEY_COLORED_AGENT, coloredAgent, cache);
    cachePutChunked_(KEY_COLORED_ADMIN, coloredAdmin, cache);

    let msg = `✅ تم التنفيذ`;
    if (copied)        msg += ` — نُسخ ${copied} صف`;
    if (skipped)       msg += ` — تم تخطي ${skipped}`;
    if (totalColored)  msg += ` — تلوين ${totalColored} صف`;
    if (lastUsedColor) msg += ` — لون ${lastUsedColor}`;
    if (!doAdminOps)   msg += ` — (وضع: الوكيل فقط)`;

    return { success:true, message: msg };
  } catch(e){
    return {success:false,message:"خطأ: "+e.message};
  }
}

/*****************************
 * القائمة + العرض
 *****************************/
function getAdminSheets(){
  const cfg = getConfig_();
  const adSS = SpreadsheetApp.openById(cfg.ADMIN_SHEET_ID);
  return adSS.getSheets().map(function(sh){ return sh.getName(); });
}
function createAdminSheet(name){
  name = String(name||'').trim();
  if(!name) return "⚠️ اكتب اسم ورقة";
  const cfg = getConfig_();
  const adSS = SpreadsheetApp.openById(cfg.ADMIN_SHEET_ID);
  if(adSS.getSheetByName(name)) return "⚠️ الورقة موجودة بالفعل";
  adSS.insertSheet(name);
  return "✅ تم إنشاء الورقة: "+name;
}

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('أداة البحث المتقدم')
      .addItem('🚀 فتح الأداة', 'showSidebar')
      .addToUi();
  } catch (_) {}
}

function showSidebar() {
  const t = HtmlService.createTemplateFromFile('Sidebar');
  t.MODE = 'sidebar';
  const html = t.evaluate().setTitle('أداة البحث').setWidth(460);
  SpreadsheetApp.getUi().showSidebar(html);
}

function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Sidebar');
  t.MODE = 'web';
  return t.evaluate()
    .setTitle('أداة البحث المتقدم')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/***** === تشغيل/إيقاف تصحيح الراتب === *****/
function setSalaryCorrectionEnabled(enabled) {
  var v = (enabled === true || enabled === '1' || enabled === 1) ? '1' : '0';
  PropertiesService.getScriptProperties().setProperty('USE_SAL_CORR', v);
  return { ok:true, enabled: v === '1' };
}
function getSalaryCorrectionEnabled() {
  var v = PropertiesService.getScriptProperties().getProperty('USE_SAL_CORR');
  return { ok:true, enabled: v === '1' };
}

/***** === Gemini test (كما هو) ===*****/
function testGemini() {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) {
      Logger.log("❌ ما في API Key مخزّن. ضيفه في Script Properties باسم GEMINI_API_KEY.");
      return;
    }

    const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=" + apiKey;

    const payload = {
      contents: [{
        parts: [{ text: "اكتب لي جملة ترحيب قصيرة باللهجة السورية" }]
      }]
    };

    const response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const result = JSON.parse(response.getContentText());
    Logger.log(result);

    const text = result?.candidates?.[0]?.content?.parts?.[0]?.text || "❌ ما في نص بالرد";
    Logger.log("✅ رد Gemini: " + text);

  } catch (e) {
    Logger.log("⚠️ خطأ: " + e.message);
  }
}

function ping(){
  return true;
}

/* ============================================================
 *      ✦✦✦  إضافات السجل والنقل (قص-لصق)  ✦✦✦
 *  (أحدث للأقدم، نقل من/إلى أي ورقة، منع التكرار، حفظ اللون)
 * ============================================================*/

/** فحص بسيط (تستخدمه الواجهة إن احتجت) */
function lgp__ping(){ return { ok:true, ts:new Date().toISOString() }; }

/** يسجّل عملية واحدة في السجل */
function logCopyOperation(targetSheetName, startAt, rowsCount, colorHex, idMaybe) {
  try {
    var props = PropertiesService.getDocumentProperties();
    var raw   = props.getProperty('COPY_LOG_V1') || '[]';
    var log;
    try { log = JSON.parse(raw); } catch(_) { log = []; }
    log.push({
      t: Date.now(),
      target: String(targetSheetName || ''),
      start:  Number(startAt || 0),
      cnt:    Number(rowsCount || 1),
      color:  String(colorHex || ''),
      id:     String(idMaybe || '')
    });
    if (log.length > 200) log = log.slice(-200);
    props.setProperty('COPY_LOG_V1', JSON.stringify(log));
  } catch(_) {}
}

/** يرجع آخر N عناصر من السجل — ترتيب: الأحدث أولاً */
function getMoveLogLatest(limit){
  var LIM = Math.max(1, Math.min(100, Number(limit || 15)));
  var props = PropertiesService.getDocumentProperties();
  var log = [];
  try { log = JSON.parse(props.getProperty('COPY_LOG_V1') || '[]'); } catch(_){}
  var last = log.slice(-LIM).reverse();
  return last.map(function(it){
    return {
      id: it.id || '',
      target: it.target || '',
      rows: it.cnt || 1,
      color: it.color || '',
      at: new Date(it.t || Date.now()).toLocaleString('ar-SY')
    };
  });
}

/** أسماء أوراق ملف الإدارة بأمان (يستخدم getAdminSheets إن وُجد) */
function getAdminSheetsSafe(){
  try {
    var arr = (typeof getAdminSheets === 'function') ? getAdminSheets() : null;
    if (Array.isArray(arr)) return arr;
  } catch(_){}
  var cfg = getConfig_();
  var adSS = SpreadsheetApp.openById(cfg.ADMIN_SHEET_ID);
  return adSS.getSheets().map(function(sh){ return sh.getName(); });
}

/** مساعد: ابحث عن ID في أي ورقة داخل ملف الإدارة (عمود A) */
function findIdInAnyAdminSheet_(ss, id){
  id = String(id||'').trim();
  if (!id) return null;
  var sheets = ss.getSheets();
  for (var s=0; s<sheets.length; s++){
    var sh = sheets[s];
    var lr = sh.getLastRow();
    if (lr < 1) continue;
    var colA = sh.getRange(1,1,lr,1).getDisplayValues();
    for (var i=0; i<lr; i++){
      if (String(colA[i][0]||'').trim() === id) {
        return { sheet: sh, row: i+1 };
      }
    }
  }
  return null;
}

/**
 * نقل كامل (قص-لصق) من/إلى وبالعكس — يمنع التكرار ويحفظ اللون
 * picks: [{id, from?}]  — لو from غير موجود، نبحث في كل الأوراق
 * targetSheetName: اسم الورقة الهدف
 * overrideHex: لون مخصّص (اختياري)
 */
function moveFromLog(picks, targetSheetName, overrideHex){
  try {
    picks = Array.isArray(picks) ? picks : [];
    targetSheetName = String(targetSheetName||'').trim();
    overrideHex     = String(overrideHex||'').trim();

    if (!picks.length)    return { ok:false, message:'لا يوجد عناصر للنقل' };
    if (!targetSheetName) return { ok:false, message:'اختر ورقة الهدف' };

    var cfg  = getConfig_();
    var ss   = SpreadsheetApp.openById(cfg.ADMIN_SHEET_ID);
    var tgSh = ss.getSheetByName(targetSheetName);
    if (!tgSh) return { ok:false, message:'ورقة الهدف غير موجودة: '+targetSheetName };

    // IDs الموجودة مسبقًا بالهدف
    var targetIdSet = (function(){
      var set = Object.create(null);
      var lr = tgSh.getLastRow();
      if (lr > 0){
        var a = tgSh.getRange(1,1,lr,1).getDisplayValues();
        for (var i=0;i<lr;i++){ var v=(a[i][0]||'').trim(); if(v) set[v]=1; }
      }
      return set;
    })();

    var moved=0, skipped=0, skippedSameSheet=0, skippedExists=0, errors=0;

    picks.forEach(function(p){
      try{
        var id   = String((p && p.id) || '').trim();
        var from = String((p && p.from) || '').trim(); // اختياري
        if (!id) { skipped++; return; }

        var srcSh, rowIdx;

        if (from){
          srcSh = ss.getSheetByName(from);
          if (!srcSh){ skipped++; return; }
          var lr  = srcSh.getLastRow(); if (lr<1){ skipped++; return; }
          var colA = srcSh.getRange(1,1,lr,1).getDisplayValues();
          rowIdx = -1;
          for (var i=0;i<lr;i++){ if (String(colA[i][0]||'').trim()===id){ rowIdx=i+1; break; } }
          if (rowIdx === -1){ skipped++; return; }
        } else {
          var hit = findIdInAnyAdminSheet_(ss, id);
          if (!hit){ skipped++; return; }
          srcSh  = hit.sheet;
          rowIdx = hit.row;
          from   = srcSh.getName();
        }

        if (from === targetSheetName){ skippedSameSheet++; return; }
        if (targetIdSet[id]) { skippedExists++; return; }

        var lastCol = srcSh.getLastColumn();
        var rngRow  = srcSh.getRange(rowIdx,1,1,lastCol);
        var vals    = rngRow.getValues()[0];

        // لون الصف الأصلي
        var srcColor = (function(){
          try {
            var rowBgs = rngRow.getBackgrounds()[0] || [];
            for (var k=0;k<rowBgs.length;k++){
              var c = (rowBgs[k]||'').toString().toLowerCase();
              if (c && c!=='#ffffff' && c!=='#fff' && c!=='transparent') return rowBgs[k];
            }
            var ca = (srcSh.getRange(rowIdx,1,1,1).getBackground()||'').toLowerCase();
            return (ca && ca!=='transparent') ? ca : '';
          } catch(e){ return ''; }
        })();

        var useColor = overrideHex || srcColor;

        // لصق في الهدف
        var destRow = tgSh.getLastRow() + 1;
        tgSh.appendRow(vals);
        if (useColor){
          var lastColTarget = tgSh.getLastColumn();
          tgSh.getRange(destRow,1,1,lastColTarget).setBackground(useColor);
        }
        targetIdSet[id] = 1;

        // حذف من المصدر (قص فعلي)
        srcSh.deleteRow(rowIdx);

        moved++;

        // سجل العملية
        try { logCopyOperation(tgSh.getName(), destRow, 1, (useColor||''), id); } catch(_){}

      } catch(e){ errors++; }
    });

    SpreadsheetApp.flush();

    var parts=[];
    parts.push('تم النقل: '+moved);
    if (skippedExists)    parts.push('موجود مسبقًا: '+skippedExists);
    if (skippedSameSheet) parts.push('نفس الورقة: '+skippedSameSheet);
    if (skipped)          parts.push('تخطي: '+skipped);
    if (errors)           parts.push('أخطاء: '+errors);

    return { ok:true, message: parts.join(' • '), moved:moved, skipped:skipped, skippedExists:skippedExists, skippedSameSheet:skippedSameSheet, errors:errors };

  } catch(e){
    return { ok:false, message:'خطأ: '+e.message };
  }
}

/** Helper to include partial HTML files in templates */
function include(name){
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}


/** Open the UI inside Sheets sidebar */
function openSidebar(){
  var ui = SpreadsheetApp.getUi();
  var html = HtmlService.createTemplateFromFile('Sidebar').evaluate()
    .setTitle('أداة البحث');
  ui.showSidebar(html);
}/** 🔒 حارس يمنع نسخ صف مكرر: 
 *   - يتحقق من وجود الـID أصلاً في ورقة الهدف
 *   - يمنع نسخ نفس الـID أكثر من مرة
 */
function alreadyCopied_(id, tgSh) {
  try {
    if (!id || !tgSh) return true; // لا تتابع إذا مافي ID أو ورقة
    id = String(id).trim();
    if (!id) return true;

    // اقرأ العمود A كله كـقيم
    const lr = tgSh.getLastRow();
    if (lr < 1) return false;

    const colA = tgSh.getRange(1,1,lr,1).getDisplayValues();
    for (let i=0; i<lr; i++){
      if ((colA[i][0]||'').trim() === id){
        return true; // ✅ موجود بالفعل
      }
    }
    return false; // مش موجود → ممكن نسخه
  } catch(e){
    return true; // أي خطأ = اعتبره موجود لتجنّب التكرار
  }
}