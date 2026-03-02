function doGet(e) {
  return ContentService.createTextOutput("API Kho Bệnh Viện đang chạy (Mode: TEXT/JSON)...").setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  var output = { success: false, msg: "Lỗi khởi tạo" };
  try {
    var jsonString = e.postData.contents;
    var params = JSON.parse(jsonString);
    var action = params.action;

    if (action === 'login') output = login(params.u, params.p);
    else if (action === 'getInitialData') output = getInitialData(params.isAdmin);
    else if (action === 'saveUser') output = saveUser(params.data);
    else if (action === 'deleteUser') output = deleteUser(params.user);
    else if (action === 'saveDMBatch') output = saveDMBatch(params.dataList);
    else if (action === 'updateDMSingle') output = updateDMSingle(params.data, params.oldKey);
    else if (action === 'deleteDM') output = deleteDM(params.key);
    else if (action === 'renameMachine') output = renameMachine(params.kho, params.oldName, params.newName);
    else if (action === 'processImport') output = processImport(params.dataList);
    else if (action === 'processExport') output = processExport(params.data);
    else if (action === 'getReport') output = getInventoryAndHistory(params.from, params.to);
    else if (action === 'changePassword') output = changePassword(params.user, params.oldPass, params.newPass);
    else output = { success: false, msg: "Action not found: " + action };

  } catch (err) {
    output = { success: false, msg: "Lỗi Server Backend: " + err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(output)).setMimeType(ContentService.MimeType.TEXT);
}

function login(u, p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TAI_KHOAN");
  var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 5).getValues();
  for (var i = 0; i < data.length; i++) {
    var user = String(data[i][0]).trim().toLowerCase();
    var pass = String(data[i][1]).trim();
    var role = String(data[i][3]).trim();
    var status = String(data[i][4]).trim();
    if (user === String(u).trim().toLowerCase() && pass === String(p).trim()) {
      if (status.toLowerCase() === 'block') return { success: false, msg: "Tài khoản bị khóa!" };
      return { success: true, name: data[i][2], user: data[i][0], role: role };
    }
  }
  return { success: false, msg: "Sai tài khoản/mật khẩu!" };
}

function getInitialData(isAdmin) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDM = ss.getSheetByName("DM");
  var dmRaw = sheetDM.getLastRow() > 1 ? sheetDM.getRange(2, 1, sheetDM.getLastRow()-1, 9).getValues() : [];
  var users = [];
  if (isAdmin) {
    var sheetUser = ss.getSheetByName("TAI_KHOAN");
    if(sheetUser.getLastRow() > 1) users = sheetUser.getRange(2, 1, sheetUser.getLastRow()-1, 5).getValues();
  }
  return { dm: dmRaw, users: users };
}

function saveUser(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("TAI_KHOAN");
  var list = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues().flat();
  var rowIndex = -1;
  for(var i=0; i<list.length; i++){ if(String(list[i]).toLowerCase() === String(data.user).toLowerCase()){ rowIndex = i + 2; break; } }
  if (rowIndex > 0) { sheet.getRange(rowIndex, 2, 1, 4).setValues([[data.pass, data.name, data.role, data.status]]); return {success: true, msg: "Đã cập nhật!"}; }
  else { sheet.appendRow([data.user, data.pass, data.name, data.role, data.status]); return {success: true, msg: "Đã thêm mới!"}; }
}

function deleteUser(u) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("TAI_KHOAN");
  var list = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues().flat();
  for(var i=0; i<list.length; i++){ if(String(list[i]).toLowerCase() === String(u).toLowerCase()){ sheet.deleteRow(i + 2); return {success: true, msg: "Đã xóa!"}; } }
  return {success: false, msg: "Không tìm thấy!"};
}

function saveDMBatch(dataList) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("DM");
  var list = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow()-1, 3).getValues() : [];
  var count = 0;
  dataList.forEach(data => {
    var exists = false;
    for(var i=0; i<list.length; i++){ if(list[i][0] == data.kho && list[i][1] == data.may && list[i][2] == data.ten) { exists = true; break; } }
    if(!exists) { sheet.appendRow([data.kho, data.may, data.ten, data.loai, data.dv, data.minR, data.minY, data.hanR, data.hanY]); list.push([data.kho, data.may, data.ten]); count++; }
  });
  return {success: true, msg: "Đã thêm " + count + " dòng."};
}

function updateDMSingle(data, oldKey) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("DM");
  var list = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow()-1, 3).getValues() : [];
  for(var i=0; i<list.length; i++) {
    var key = list[i][0] + "|" + list[i][1] + "|" + list[i][2];
    if (key === oldKey) {
      sheet.getRange(i + 2, 1, 1, 9).setValues([[data.kho, data.may, data.ten, data.loai, data.dv, data.minR, data.minY, data.hanR, data.hanY]]);
      return {success: true, msg: "Đã cập nhật!"};
    }
  }
  return {success: false, msg: "Lỗi: Không tìm thấy dòng cũ."};
}

function deleteDM(keyDel) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("DM");
  var list = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow()-1, 3).getValues() : [];
  for(var i=0; i<list.length; i++) { if ((list[i][0]+"|"+list[i][1]+"|"+list[i][2]) === keyDel) { sheet.deleteRow(i + 2); return {success: true, msg: "Đã xóa!"}; } }
  return {success: false, msg: "Lỗi!"};
}

function renameMachine(kho, oldName, newName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("DM");
  var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 2).getValues(); var count = 0;
  for(var i=0; i<data.length; i++) { if(data[i][0] == kho && data[i][1] == oldName) { sheet.getRange(i + 2, 2).setValue(newName); count++; } }
  return {success: true, msg: "Đã cập nhật " + count + " dòng."};
}

function processImport(dataList) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NHAP"); var ts = new Date();
  var tonKhoUpdates = [];
  dataList.forEach(data => {
    var lot2 = data.loaiNhap === 'R1R2' ? data.lotR2 : ""; var hsd2 = data.loaiNhap === 'R1R2' ? data.hsdR2 : ""; var sl2 = data.loaiNhap === 'R1R2' ? data.slR2 : "";
    sheet.appendRow([ts, new Date(data.ngayNhap), data.nguoi, data.kho, data.may, data.tenHC, data.lotR1, new Date(data.hsdR1), data.slR1, lot2, hsd2?new Date(hsd2):"", sl2, data.donVi]);

    tonKhoUpdates.push({
      kho: data.kho,
      may: data.may,
      ten: data.tenHC,
      part: 'R1',
      lot: data.lotR1,
      hsd: data.hsdR1,
      qty: toSafeNumber_(data.slR1),
      donVi: data.donVi
    });

    if (data.loaiNhap === 'R1R2' && data.lotR2) {
      tonKhoUpdates.push({
        kho: data.kho,
        may: data.may,
        ten: data.tenHC,
        part: 'R2',
        lot: data.lotR2,
        hsd: data.hsdR2,
        qty: toSafeNumber_(data.slR2),
        donVi: data.donVi
      });
    }
  });

  applyTonKhoUpdates_(tonKhoUpdates);
  return {success: true, msg: "Nhập kho thành công " + dataList.length + " phiếu!"};
}

function processExport(data) {
  var today = new Date(); today.setHours(0,0,0,0); var exportDate = new Date(data.ngayXuat); exportDate.setHours(0,0,0,0);
  if (exportDate < today && data.lyDo !== "Quên không xuất kho") return {success: false, msg: "Lỗi: Ngày quá khứ sai lý do."};
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("XUAT");
  sheet.appendRow([new Date(), new Date(data.ngayXuat), data.nguoi, data.kho, data.may, data.tenHC, data.lotR1, data.slR1, data.lotR2, data.slR2, data.lyDo]);

  var tonKhoUpdates = [];
  if (data.lotR1) {
    tonKhoUpdates.push({
      kho: data.kho,
      may: data.may,
      ten: data.tenHC,
      part: 'R1',
      lot: data.lotR1,
      hsd: "",
      qty: -toSafeNumber_(data.slR1),
      donVi: ""
    });
  }
  if (data.lotR2) {
    tonKhoUpdates.push({
      kho: data.kho,
      may: data.may,
      ten: data.tenHC,
      part: 'R2',
      lot: data.lotR2,
      hsd: "",
      qty: -toSafeNumber_(data.slR2),
      donVi: ""
    });
  }

  applyTonKhoUpdates_(tonKhoUpdates);
  return {success: true, msg: "Xuất kho thành công!"};
}

function changePassword(user, oldP, newP) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("TAI_KHOAN");
  var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 2).getValues();
  for(var i=0; i<data.length; i++){
    if(String(data[i][0]).toLowerCase() === String(user).toLowerCase()){
      if(String(data[i][1])===String(oldP)){
        sheet.getRange(i+2,2).setValue(newP);
        return {success: true, msg: "SUCCESS"};
      } else return {success: false, msg: "Mật khẩu cũ sai!"};
    }
  }
  return {success: false, msg: "Lỗi!"};
}

function toSafeNumber_(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  var str = String(v).trim();
  if (!str) return 0;
  str = str.replace(/\s+/g, '').replace(/,/g, '.');
  var n = Number(str);
  return isNaN(n) ? 0 : n;
}

function normalizeStr_(s) {
  return String(s == null ? "" : s).trim();
}

function normalizeLotKey_(lot) {
  return normalizeStr_(lot).toUpperCase();
}

function parseDateSafe_(d) {
  if (d === null || d === undefined || d === "") return null;
  if (d instanceof Date) return isNaN(d.getTime()) ? null : new Date(d.getTime());

  var direct = new Date(d);
  if (!isNaN(direct.getTime())) return direct;

  var str = String(d).trim();
  var m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
  if (m) {
    var day = Number(m[1]), mon = Number(m[2]) - 1, year = Number(m[3]);
    var hh = Number(m[4] || 0), mm = Number(m[5] || 0), ss = Number(m[6] || 0);
    var parsed = new Date(year, mon, day, hh, mm, ss);
    return isNaN(parsed.getTime()) ? null : parsed;
  }
  return null;
}

function getOrCreateTonKhoSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TONKHO");
  if (!sheet) {
    sheet = ss.insertSheet("TONKHO");
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Kho", "May", "Ten_HC", "Loai", "Lot", "Han", "SL", "Don_Vi", "Updated_At"]);
  }
  return sheet;
}

function applyTonKhoUpdates_(updates) {
  if (!updates || updates.length === 0) return;

  var sheet = getOrCreateTonKhoSheet_();
  var rows = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow()-1, 9).getValues() : [];
  var map = {};

  function rowKey_(kho, may, ten, loai, lot) {
    return normalizeStr_(kho) + "|" + normalizeStr_(may) + "|" + normalizeStr_(ten) + "|" + normalizeStr_(loai) + "|" + normalizeLotKey_(lot);
  }

  for (var i = 0; i < rows.length; i++) {
    map[rowKey_(rows[i][0], rows[i][1], rows[i][2], rows[i][3], rows[i][4])] = { idx: i, row: rows[i] };
  }

  var now = new Date();
  var touched = {};
  var appendRows = [];

  updates.forEach(function(u) {
    var qty = toSafeNumber_(u.qty);
    var lot = normalizeStr_(u.lot);
    if (!lot || qty === 0) return;

    var key = rowKey_(u.kho, u.may, u.ten, u.part, lot);
    var found = map[key];

    if (found) {
      var row = found.row;
      row[6] = toSafeNumber_(row[6]) + qty;

      var hsdObj = parseDateSafe_(u.hsd);
      if (hsdObj) row[5] = hsdObj;
      if (normalizeStr_(u.donVi)) row[7] = normalizeStr_(u.donVi);
      row[8] = now;

      touched[found.idx] = row;
    } else {
      var hsdObj2 = parseDateSafe_(u.hsd);
      appendRows.push([
        normalizeStr_(u.kho),
        normalizeStr_(u.may),
        normalizeStr_(u.ten),
        normalizeStr_(u.part),
        lot,
        hsdObj2 || "",
        qty,
        normalizeStr_(u.donVi),
        now
      ]);
    }
  });

  Object.keys(touched).forEach(function(idxStr) {
    var idx = Number(idxStr);
    sheet.getRange(idx + 2, 1, 1, 9).setValues([touched[idx]]);
  });

  if (appendRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, 9).setValues(appendRows);
  }
}

function rebuildTonKhoSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNhap = ss.getSheetByName("NHAP");
  var sheetXuat = ss.getSheetByName("XUAT");
  var tonSheet = getOrCreateTonKhoSheet_();

  var nhapData = (sheetNhap && sheetNhap.getLastRow() > 1) ? sheetNhap.getRange(2, 1, sheetNhap.getLastRow()-1, 13).getValues() : [];
  var xuatData = (sheetXuat && sheetXuat.getLastRow() > 1) ? sheetXuat.getRange(2, 1, sheetXuat.getLastRow()-1, 11).getValues() : [];

  var map = {};
  function key(kho, may, ten, loai, lot) {
    return normalizeStr_(kho) + "|" + normalizeStr_(may) + "|" + normalizeStr_(ten) + "|" + normalizeStr_(loai) + "|" + normalizeLotKey_(lot);
  }

  function add(kho, may, ten, loai, lot, hsd, qty, dv) {
    var lotStr = normalizeStr_(lot);
    var q = toSafeNumber_(qty);
    if (!lotStr || q === 0) return;
    var k = key(kho, may, ten, loai, lotStr);
    if (!map[k]) {
      var hsdObj = parseDateSafe_(hsd);
      map[k] = {
        kho: normalizeStr_(kho),
        may: normalizeStr_(may),
        ten: normalizeStr_(ten),
        loai: normalizeStr_(loai),
        lot: lotStr,
        han: hsdObj || "",
        sl: 0,
        dv: normalizeStr_(dv)
      };
    }
    map[k].sl += q;
    var hsdObj2 = parseDateSafe_(hsd);
    if (hsdObj2) map[k].han = hsdObj2;
    if (normalizeStr_(dv)) map[k].dv = normalizeStr_(dv);
  }

  nhapData.forEach(function(r) {
    add(r[3], r[4], r[5], 'R1', r[6], r[7], r[8], r[12]);
    add(r[3], r[4], r[5], 'R2', r[9], r[10], r[11], r[12]);
  });
  xuatData.forEach(function(r) {
    add(r[3], r[4], r[5], 'R1', r[6], "", -toSafeNumber_(r[7]), "");
    add(r[3], r[4], r[5], 'R2', r[8], "", -toSafeNumber_(r[9]), "");
  });

  if (tonSheet.getLastRow() > 1) tonSheet.getRange(2, 1, tonSheet.getLastRow()-1, 9).clearContent();

  var out = [];
  var now = new Date();
  Object.keys(map).forEach(function(k) {
    var it = map[k];
    out.push([it.kho, it.may, it.ten, it.loai, it.lot, it.han, it.sl, it.dv, now]);
  });

  if (out.length > 0) tonSheet.getRange(2, 1, out.length, 9).setValues(out);
}

function getInventoryAndHistory(fromStr, toStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetNhap = ss.getSheetByName("NHAP");
    var sheetXuat = ss.getSheetByName("XUAT");
    var tonSheet = getOrCreateTonKhoSheet_();

    var nhapData = (sheetNhap && sheetNhap.getLastRow() > 1) ? sheetNhap.getRange(2, 1, sheetNhap.getLastRow()-1, 13).getValues() : [];
    var xuatData = (sheetXuat && sheetXuat.getLastRow() > 1) ? sheetXuat.getRange(2, 1, sheetXuat.getLastRow()-1, 11).getValues() : [];
    if (tonSheet.getLastRow() <= 1) {
      rebuildTonKhoSheet_();
    }
    var tonData = tonSheet.getLastRow() > 1 ? tonSheet.getRange(2, 1, tonSheet.getLastRow()-1, 9).getValues() : [];

    var inventory = {};
    var timeline = [];

    function ensureItem(key) {
      if (!inventory[key]) inventory[key] = { R1: {}, R2: {} };
      return inventory[key];
    }

    function applyLotDelta(stockPart, lotRaw, hsdRaw, qtyDelta) {
      var lk = normalizeLotKey_(lotRaw);
      if (!lk) return;

      if (!stockPart[lk]) {
        var hsdObj = parseDateSafe_(hsdRaw);
        stockPart[lk] = {
          lot: normalizeStr_(lotRaw),
          hsd: hsdObj ? formatDate(hsdObj) : "",
          rawHsd: hsdObj ? hsdObj.getTime() : 0,
          sl: 0
        };
      }

      if (!stockPart[lk].lot) stockPart[lk].lot = normalizeStr_(lotRaw);
      if (!stockPart[lk].rawHsd) {
        var hsdObj2 = parseDateSafe_(hsdRaw);
        if (hsdObj2) {
          stockPart[lk].rawHsd = hsdObj2.getTime();
          stockPart[lk].hsd = formatDate(hsdObj2);
        }
      }

      stockPart[lk].sl += qtyDelta;
    }

    tonData.forEach(function(r) {
      var kho = normalizeStr_(r[0]);
      var may = normalizeStr_(r[1]);
      var ten = normalizeStr_(r[2]);
      var loai = normalizeStr_(r[3]).toUpperCase();
      var lot = normalizeStr_(r[4]);
      var hsd = r[5];
      var sl = toSafeNumber_(r[6]);
      if (!kho || !ten || !lot) return;

      var key = kho + "|" + may + "|" + ten;
      var item = ensureItem(key);
      if (loai === 'R2') applyLotDelta(item.R2, lot, hsd, sl);
      else applyLotDelta(item.R1, lot, hsd, sl);
    });

    nhapData.forEach(function(r) {
      var kho = normalizeStr_(r[3]);
      var may = normalizeStr_(r[4]);
      var ten = normalizeStr_(r[5]);
      if (!kho || !ten) return;

      var lotR1Raw = normalizeStr_(r[6]);
      var lotR2Raw = normalizeStr_(r[9]);
      var slR1 = toSafeNumber_(r[8]);
      var slR2 = toSafeNumber_(r[11]);

      var ioType = (lotR1Raw && lotR2Raw) ? 'Cả R1,R2' : (lotR2Raw ? 'R2' : 'R1');

      var ts = parseDateSafe_(r[0]) || new Date();
      var dDoc = parseDateSafe_(r[1]) || ts;

      timeline.push({
        type: 'NHẬP',
        action: 'NHẬP',
        sortTime: ts.getTime(),
        timestampStr: formatDateLong(ts),
        dateDocStr: formatDate(dDoc),
        user: r[2], kho: kho, may: may, ten: ten,
        r1_lot: lotR1Raw, r1_sl: slR1,
        r2_lot: lotR2Raw, r2_sl: slR2,
        ioType: ioType,
        reason: "Nhập mới"
      });
    });

    xuatData.forEach(function(r) {
      var kho = normalizeStr_(r[3]);
      var may = normalizeStr_(r[4]);
      var ten = normalizeStr_(r[5]);
      if (!kho || !ten) return;

      var lotR1Raw = normalizeStr_(r[6]);
      var lotR2Raw = normalizeStr_(r[8]);
      var slR1 = toSafeNumber_(r[7]);
      var slR2 = toSafeNumber_(r[9]);

      var ioType = (lotR1Raw && lotR2Raw) ? 'Cả R1,R2' : (lotR2Raw ? 'R2' : 'R1');

      var ts = parseDateSafe_(r[0]) || new Date();
      var dDoc = parseDateSafe_(r[1]) || ts;

      timeline.push({
        type: 'XUẤT',
        action: 'XUẤT',
        sortTime: ts.getTime(),
        timestampStr: formatDateLong(ts),
        dateDocStr: formatDate(dDoc),
        user: r[2], kho: kho, may: may, ten: ten,
        r1_lot: lotR1Raw, r1_sl: slR1,
        r2_lot: lotR2Raw, r2_sl: slR2,
        ioType: ioType,
        reason: r[10]
      });
    });

    var timelineAsc = timeline.slice().sort(function(a,b){ return a.sortTime - b.sortTime; });
    var runningTracker = {};

    var useFilter = (fromStr && toStr);
    var sDate = useFilter ? parseDateSafe_(fromStr) : null;
    var eDate = useFilter ? parseDateSafe_(toStr) : null;
    var sTs = sDate ? sDate.setHours(0,0,0,0) : 0;
    var eTs = eDate ? eDate.setHours(23,59,59,999) : 0;

    var result = [];

    timelineAsc.forEach(function(item) {
      var d = [];
      var base = item.kho + "|" + item.may + "|" + item.ten;
      var isNhap = (item.action === 'NHẬP');

      d.push(`<b>Loại:</b> ${item.ioType || 'R1'}`);

      if(item.r1_lot){
        var k1 = base + "|R1|" + normalizeLotKey_(item.r1_lot);
        if(runningTracker[k1] == null) runningTracker[k1] = 0;
        runningTracker[k1] += isNhap ? item.r1_sl : -item.r1_sl;

        var style1 = isNhap ? 'text-success fw-bold' : 'text-danger fw-bold';
        var sign1 = isNhap ? '+' : '-';
        d.push(`<b>R1:</b> ${item.r1_lot} <span class="${style1}">(${sign1}${item.r1_sl})</span> <span class="text-muted small">[Tồn: ${runningTracker[k1]}]</span>`);
      }

      if(item.r2_lot){
        var k2 = base + "|R2|" + normalizeLotKey_(item.r2_lot);
        if(runningTracker[k2] == null) runningTracker[k2] = 0;
        runningTracker[k2] += isNhap ? item.r2_sl : -item.r2_sl;

        var style2 = isNhap ? 'text-success fw-bold' : 'text-danger fw-bold';
        var sign2 = isNhap ? '+' : '-';
        d.push(`<b>R2:</b> ${item.r2_lot} <span class="${style2}">(${sign2}${item.r2_sl})</span> <span class="text-muted small">[Tồn: ${runningTracker[k2]}]</span>`);
      }

      item.detailHtml = d.join("<br>");

      if(!useFilter || (item.sortTime >= sTs && item.sortTime <= eTs)){
        if(d.length > 0 && item.kho) result.push(item);
      }
    });

    result.sort(function(a,b){ return b.sortTime - a.sortTime; });

    var stockOut = {};
    Object.keys(inventory).forEach(function(mainKey) {
      stockOut[mainKey] = { R1: {}, R2: {} };
      ['R1', 'R2'].forEach(function(part) {
        Object.keys(inventory[mainKey][part]).forEach(function(lk) {
          var obj = inventory[mainKey][part][lk];
          stockOut[mainKey][part][obj.lot || lk] = {
            hsd: obj.hsd || "",
            rawHsd: obj.rawHsd || 0,
            sl: obj.sl
          };
        });
      });
    });

    return { stock: stockOut, history: useFilter ? result : result.slice(0,50) };

  } catch (e) {
    return {stock:{}, history:[], error: "Lỗi tính toán Backend: " + e.toString()};
  }
}

function formatDate(d) {
  if (!d) return "";
  try {
    var date = (d instanceof Date) ? d : new Date(d);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
  } catch(e){ return ""; }
}

function formatDateLong(d) {
  if (!d) return "";
  try {
    var date = (d instanceof Date) ? d : new Date(d);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  } catch(e){ return ""; }
}
