const DATA_SHEET = 'Data_CTKM';
const TIMELINE_SHEET = 'Timeline';
const MASTERDATA_SHEET = 'MasterData';
const SALES_SHEET = 'DATA_MASTER';
const GUIDELINE_SHEET = 'Guideline_CTKM';

function doGet(e) {
  if (!e || !e.parameter || !e.parameter.action) {
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Quản Lý CTKM')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  try {
    switch (e.parameter.action) {
      case 'list':
        return jsonSuccess(listCtkm());
      case 'masterdata':
        return jsonSuccess(loadMasterData());
      case 'guideline':
        return jsonSuccess(loadGuidelineData());
      default:
        return jsonSuccess([]);
    }
  } catch (err) {
    return jsonError(err.toString());
  }
}

function doPost(e) {
  const payload = JSON.parse(e.postData.contents || '{}');
  const action = payload.action || 'save';

  try {
    if (action === 'save')         return jsonSuccess(saveBatchCtkm(payload.data || {}));
    if (action === 'save_pdf')     return jsonSuccess(savePdfToDrive(payload.data));
    if (action === 'save_timeline_pdf') return jsonSuccess(saveTimelinePdf(payload.data));
    if (action === 'report')       return jsonSuccess(getCtkmReport(payload.data.name));
    if (action === 'update_timeline') return jsonSuccess(updateTimeline(payload.data));
    if (action === 'save_guideline') return jsonSuccess(saveGuideline(payload.data));
    if (action === 'delete_guideline') return jsonSuccess(deleteGuideline(payload.data));
    return jsonError('Unknown action');
  } catch (err) {
    return jsonError(err.toString());
  }
}

function cleanStr(str) {
  return String(str || "").toUpperCase().replace(/[^A-Z0-9]/g, '');
}

// --- 1. BÁO CÁO DOANH SỐ (CHI TIẾT SKU) ---
function getCtkmReport(ctkmName) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const sSales   = ss.getSheetByName(SALES_SHEET);
  const sTimeline= ss.getSheetByName(TIMELINE_SHEET);
  const sPromo   = ss.getSheetByName(DATA_SHEET);

  if (!sSales) return { error: "Không tìm thấy sheet " + SALES_SHEET };

  const nameTarget = String(ctkmName).toUpperCase().trim();

  // Lấy thời gian CTKM từ sheet Timeline
  const tlVals = sTimeline.getDataRange().getValues();
  let start = null, end = null;

  for (let i = 1; i < tlVals.length; i++) {
    if (String(tlVals[i][1]).toUpperCase() === nameTarget) {
      const range = parseCustomDateRange(tlVals[i][2]);
      start = range.startDate ? new Date(range.startDate) : null;
      end   = range.endDate   ? new Date(range.endDate)   : null;
      break;
    }
  }

  if (!start) return { error: "Không tìm thấy thời gian CTKM!" };

  start.setHours(0, 0, 0, 0);
  end.setHours(23, 59, 59, 999);

  // Xác định hệ thống (system) và danh sách SKU thuộc CTKM này
  const prVals = sPromo.getDataRange().getValues();
  let targetSys = "";
  const validSkus = new Set(); // Set chứa các SKU thuộc CTKM này

  for (let i = 1; i < prVals.length; i++) {
    if (String(prVals[i][5]).toUpperCase().trim() === nameTarget) {
      const sys = cleanStr(prVals[i][4]);
      const sku = cleanStr(prVals[i][3]); // Cột D = SKU (index 3)
      
      if (!targetSys) targetSys = sys; // Lấy hệ thống từ dòng đầu tiên
      if (sku) validSkus.add(sku); // Thêm SKU vào danh sách
    }
  }

  if (!targetSys || validSkus.size === 0) {
    return { 
      name: ctkmName, 
      system: targetSys || "",
      qty: 0, 
      revenue: 0, 
      details: [],
      error: validSkus.size === 0 ? "Không tìm thấy SKU nào trong CTKM này!" : ""
    };
  }

  const lastRow = sSales.getLastRow();
  if (lastRow < 2) {
    return { name: ctkmName, system: targetSys, qty: 0, revenue: 0, details: [] };
  }

  const salesVals = sSales.getRange(1, 1, lastRow, 16).getValues();

  let totalQty = 0, totalRev = 0;
  const detailsMap = {};

  const COL_DATE = 1;
  const COL_QTY  = 4;
  const COL_REV  = 7;
  const COL_SYS  = 12;
  const COL_SKU  = 15;

  for (let i = 1; i < salesVals.length; i++) {
    const row = salesVals[i];

    // Xử lý ngày
    let d = row[COL_DATE], saleDate = null;
    if (d instanceof Date) {
      saleDate = d;
    } else if (typeof d === 'string') {
      const p = d.trim().split('/');
      if (p.length === 3) {
        saleDate = new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
      }
    }

    if (!saleDate || isNaN(saleDate.getTime())) continue;

    const sys = cleanStr(row[COL_SYS]);
    const skuRaw = cleanStr(row[COL_SKU] || "");

    // CHỈ tính doanh số của các SKU thuộc CTKM này, trong hệ thống này, trong khoảng thời gian này
    if (saleDate >= start && saleDate <= end && 
        sys.includes(targetSys) && 
        validSkus.has(skuRaw)) {
      const q = Number(row[COL_QTY]) || 0;
      const r = Number(row[COL_REV]) || 0;

      // Dùng SKU gốc (chưa clean) để hiển thị
      const skuDisplay = String(row[COL_SKU] || "Unknown");
      if (!detailsMap[skuDisplay]) detailsMap[skuDisplay] = { qty: 0, rev: 0 };
      detailsMap[skuDisplay].qty += q;
      detailsMap[skuDisplay].rev += r;

      totalQty += q;
      totalRev += r;
    }
  }

  const details = Object.keys(detailsMap).map(k => ({
    sku: k,
    qty: detailsMap[k].qty,
    rev: detailsMap[k].rev
  }));

  const tz = Session.getScriptTimeZone();

  return {
    name: ctkmName,
    system: targetSys,
    start: Utilities.formatDate(start, tz, "dd/MM/yyyy"),
    end:   Utilities.formatDate(end,   tz, "dd/MM/yyyy"),
    qty: totalQty,
    revenue: totalRev,
    details: details
  };
}

// --- 2. LƯU BATCH ---
function saveBatchCtkm(data) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const dS  = ss.getSheetByName(DATA_SHEET);
  const tS  = ss.getSheetByName(TIMELINE_SHEET);

  const common = data.common;
  const items  = data.items;

  const name = (common.name || '').trim();

  let sD = null, eD = null;
  let rd = (common.date_range || '').replace(' đến ', ' to ');

  try {
    if (rd.includes(' to ')) {
      const p = rd.split(' to ');
      sD = parseVNDate(p[0]);
      eD = parseVNDate(p[1]);
    } else if (rd) {
      sD = parseVNDate(rd);
      eD = sD;
    }
  } catch (e) {}

  if (!sD) return { error: "Lỗi ngày tháng!" };

  const off = parseInt(common.delivery_offset || 0, 10);

  const tRow = [
    common.system,
    name,
    formatTimeRange(sD, eD),
    formatTimeRange(new Date(sD.getTime() - off * 86400000), eD)
  ];

  const tVals = tS.getDataRange().getValues();
  let tIdx = -1;

  for (let i = 1; i < tVals.length; i++) {
    if (String(tVals[i][1]).toUpperCase() === name.toUpperCase()) {
      tIdx = i + 1;
      break;
    }
  }

  if (tIdx > 0) tS.getRange(tIdx, 1, 1, tRow.length).setValues([tRow]);
  else          tS.appendRow(tRow);

  const dVals = dS.getDataRange().getValues();

  items.forEach(item => {
    let promoText = item.promotion;

    if (promoText && !isNaN(promoText)) promoText += '%';
    if (promoText && !String(promoText).toLowerCase().match(/giảm|tặng/)) {
      promoText = 'Giảm ' + promoText;
    }

    const row = [
      common.month ? `THÁNG ${common.month}` : '',
      common.year,
      common.brand,
      item.sku,
      common.system,
      name,
      promoText,
      item.gift,
      common.note
    ];

    let rowIdx = -1;
    for (let i = 1; i < dVals.length; i++) {
      if (
        cleanStr(dVals[i][5]) === cleanStr(name) &&
        cleanStr(dVals[i][3]) === cleanStr(item.sku) &&
        cleanStr(dVals[i][4]) === cleanStr(common.system)
      ) {
        rowIdx = i + 1;
        break;
      }
    }

    if (rowIdx > 0) dS.getRange(rowIdx, 1, 1, row.length).setValues([row]);
    else            dS.appendRow(row);
  });

  return { message: `Đã lưu ${items.length} mã!` };
}

// --- 3. UPDATE TIMELINE (KÉO THẢ) ---
function updateTimeline(data) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const dS  = ss.getSheetByName(DATA_SHEET);
  const tS  = ss.getSheetByName(TIMELINE_SHEET);

  const oldCtkmName = String(data.oldCtkmName || '').trim();
  const newCtkmName = String(data.newCtkmName || '').trim();
  const newStartDate = data.newStartDate ? new Date(data.newStartDate) : null;
  const newEndDate = data.newEndDate ? new Date(data.newEndDate) : null;
  const system = String(data.system || '').trim();
  const sku = String(data.sku || '').trim();

  if (!oldCtkmName || !newCtkmName || !newStartDate || !newEndDate) {
    return { error: "Thiếu thông tin!" };
  }

  // 1. Update Timeline sheet
  const tVals = tS.getDataRange().getValues();
  let tIdx = -1;
  for (let i = 1; i < tVals.length; i++) {
    if (String(tVals[i][1]).trim().toUpperCase() === newCtkmName.toUpperCase()) {
      tIdx = i + 1;
      break;
    }
  }

  const dateRangeStr = formatTimeRange(newStartDate, newEndDate);
  const deliveryOffset = parseInt(data.deliveryOffset || 0, 10);
  const deliveryDateStr = formatTimeRange(
    new Date(newStartDate.getTime() - deliveryOffset * 86400000),
    newEndDate
  );

  if (tIdx > 0) {
    tS.getRange(tIdx, 1, 1, 4).setValues([[
      system,
      newCtkmName,
      dateRangeStr,
      deliveryDateStr
    ]]);
  } else {
    tS.appendRow([system, newCtkmName, dateRangeStr, deliveryDateStr]);
  }

  // 2. Update Data_CTKM sheet: tìm dòng có oldCtkmName + system + sku, đổi tên CTKM
  const dVals = dS.getDataRange().getValues();
  let updated = 0;

  for (let i = 1; i < dVals.length; i++) {
    if (
      cleanStr(dVals[i][5]) === cleanStr(oldCtkmName) &&
      cleanStr(dVals[i][4]) === cleanStr(system) &&
      cleanStr(dVals[i][3]) === cleanStr(sku)
    ) {
      dS.getRange(i + 1, 6).setValue(newCtkmName); // Cột F = Tên CTKM
      updated++;
    }
  }

  return { message: `Đã cập nhật ${updated} dòng!` };
}

// --- UTILS LIST & MASTERDATA ---
function listCtkm() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const dS  = ss.getSheetByName(DATA_SHEET);
  const tS  = ss.getSheetByName(TIMELINE_SHEET);

  const tMap = {};
  const tV   = tS.getDataRange().getValues();

  for (let i = 1; i < tV.length; i++) {
    if (tV[i][1]) {
      tMap[String(tV[i][1]).trim().toUpperCase()] = {
        d:  tV[i][2],
        dl: tV[i][3]
      };
    }
  }

  const res = [];
  const dV  = dS.getDataRange().getValues();

  for (let i = 1; i < dV.length; i++) {
    const r = dV[i];
    if (!r[5]) continue;

    const t = tMap[String(r[5]).trim().toUpperCase()] || {};
    const dates = parseCustomDateRange(t.d);

    res.push({
      name:       r[5],
      month:      r[0],
      year:       r[1],
      brand:      r[2],
      sku:        r[3],
      system:     r[4],
      promotion:  r[6],
      gift:       r[7],
      note:       r[8],
      start_date: dates.startDate,
      end_date:   dates.endDate,
      delivery_str: t.dl
    });
  }

  return res;
}

function loadMasterData() {
  const v = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(MASTERDATA_SHEET)
    .getDataRange()
    .getValues();

  const h = {};

  for (let i = 1; i < v.length; i++) {
    const s = String(v[i][2] || '').trim();  // System
    const b = String(v[i][0] || '').trim();  // Brand

    if (!s || !b) continue;

    if (!h[s])      h[s] = {};
    if (!h[s][b])   h[s][b] = [];

    const skuVal   = String(v[i][1]);
    const priceVal = v[i][5];
    const nameVal  = v[i][3];
    const bcVal    = v[i][6];

    // Mỗi hệ thống + brand chỉ giữ 1 dòng cho mỗi SKU, chọn giá lớn nhất
    const skuKey = cleanStr(skuVal);
    const list   = h[s][b];
    const existed = list.find(x => cleanStr(x.sku) === skuKey);

    if (!existed) {
      list.push({
        sku:     skuVal,
        price:   priceVal,
        name:    nameVal,
        barcode: bcVal
      });
    } else {
      const oldPrice = Number(existed.price || 0);
      const newPrice = Number(priceVal || 0);
      if (newPrice > oldPrice) {
        existed.price   = priceVal;
        existed.name    = nameVal;
        existed.barcode = bcVal;
      }
    }
  }

  return h;
}

// --- LƯU PDF ---
/**
 * Lưu PDF vào Google Drive
 * Sử dụng HTML với CSS để tạo PDF chuyên nghiệp
 */
function savePdfToDrive(data) {
  const baseName = `VietLien_CTKM_${data.system}_${data.ctkmName}`;
  const files    = DriveApp.searchFiles(`title contains '${baseName}' and trashed=false`);

  let count = 0;
  while (files.hasNext()) {
    files.next();
    count++;
  }

  const finalName = `${baseName}_${String(count + 1).padStart(2, '0')}.pdf`;

  const html = createPdfHtml(data);
  const blob = Utilities.newBlob(html, MimeType.HTML, finalName);
  const file = DriveApp.createFile(blob.getAs(MimeType.PDF)).setName(finalName);

  return { message: 'Đã lưu file: ' + finalName, url: file.getUrl() };
}

/**
 * Setup Print Settings cho PDF
 * Áp dụng các tham số in ấn chuyên nghiệp
 */
function setupPrintSettings() {
  return {
    size: 'A4',
    portrait: false, // Landscape
    fitw: true,      // Fit to width
    gridlines: false, // Tắt đường lưới
    top_margin: 0.5,
    bottom_margin: 0.5,
    left_margin: 0.5,
    right_margin: 0.5
  };
}

/**
 * Tạo HTML cho PDF với format chuyên nghiệp
 * Áp dụng các quy định về layout, typography và spacing
 */
function createPdfHtml(d) {
  // Helper: Format số với dấu phân cách hàng ngàn (#,##0)
  const formatNumber = (num) => {
    if (!num || num === 0) return '0';
    return String(num).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  };

  // Apply formatting settings
  const printSettings = setupPrintSettings();

  let rows = '';
  d.items.forEach((item, i) => {
    // Sử dụng giá trị thô (priceVal, finalPriceVal, qtyVal, amountVal) nếu có, nếu không thì parse từ giá trị đã format
    const priceVal = item.priceVal !== undefined ? item.priceVal : (typeof item.price === 'string' ? parseFloat(String(item.price).replace(/\./g, '')) : (item.price || 0));
    const priceKMVal = item.finalPriceVal !== undefined ? item.finalPriceVal : (typeof item.priceKM === 'string' ? parseFloat(String(item.priceKM).replace(/\./g, '')) : (item.priceKM || 0));
    const qtyVal = item.qtyVal !== undefined ? item.qtyVal : (typeof item.qty === 'string' ? parseFloat(String(item.qty).replace(/\./g, '')) : (item.qty || 0));
    const amountVal = item.amountVal !== undefined ? item.amountVal : (typeof item.amount === 'string' ? parseFloat(String(item.amount).replace(/\./g, '')) : (item.amount || 0));
    
    rows += `<tr style="min-height:25px;height:auto">
      <td style="border:1px solid #666666;text-align:center;padding:8px 6px;font-size:11px;vertical-align:middle;min-height:25px">${i + 1}</td>
      <td style="border:1px solid #666666;text-align:center;font-weight:bold;padding:8px 6px;font-size:11px;vertical-align:middle;min-height:25px">${String(item.barcode || '').trim()}</td>
      <td style="border:1px solid #666666;text-align:left;padding:8px 10px;font-size:11px;vertical-align:middle;word-wrap:break-word;white-space:normal;width:250px;min-height:25px">${String(item.name || '').trim()}</td>
      <td style="border:1px solid #666666;text-align:center;padding:8px 6px;font-size:11px;vertical-align:middle;min-height:25px">${formatNumber(priceVal)}</td>
      <td style="border:1px solid #666666;text-align:center;padding:8px 6px;font-size:11px;vertical-align:middle;min-height:25px">${String(item.promo || '').trim()}</td>
      <td style="border:1px solid #666666;text-align:left;color:red;font-size:10px;padding:8px 6px;vertical-align:middle;word-wrap:break-word;min-height:25px">${String(item.gift || '').trim()}</td>
      <td style="border:1px solid #666666;text-align:right;font-weight:bold;padding:8px 6px;font-size:11px;vertical-align:middle;min-height:25px">${formatNumber(priceKMVal)}</td>
      <td style="border:1px solid #666666;text-align:center;padding:8px 6px;font-size:11px;vertical-align:middle;min-height:25px">${formatNumber(qtyVal)}</td>
      <td style="border:1px solid #666666;text-align:right;font-weight:bold;padding:8px 6px;font-size:11px;vertical-align:middle;min-height:25px">${formatNumber(amountVal)}</td>
      <td style="border:1px solid #666666;text-align:left;padding:8px 6px;font-size:11px;vertical-align:middle;word-wrap:break-word;min-height:25px">${String(item.note || '').trim()}</td>
    </tr>`;
  });

  const today = new Date();
  const day = String(today.getDate()).padStart(2, '0');
  const month = String(today.getMonth() + 1).padStart(2, '0');
  const year = today.getFullYear();

  return `
    <html>
      <head>
        <meta charset="UTF-8">
        <style>
          @page {
            size: A4 landscape;
            margin: 0.5cm; /* top_margin=0.5, bottom_margin=0.5, left_margin=0.5, right_margin=0.5 */
          }
          * {
            box-sizing: border-box;
          }
          body {
            font-family: 'Times New Roman', serif;
            font-size: 12px;
            line-height: 1.4;
            color: #000;
            margin: 0;
            padding: 0;
          }
          /* 1. HEADER & PHÂN CẤP */
          .header {
            margin-bottom: 15px;
            text-align: left;
          }
          .company-name {
            font-weight: normal;
            font-size: 10px;
            margin-bottom: 3px;
            text-transform: none;
          }
          .company-info {
            font-size: 10px;
            font-weight: normal;
            margin-bottom: 10px;
            line-height: 1.5;
          }
          .title {
            text-align: center;
            font-size: 18px;
            font-weight: bold;
            text-transform: uppercase;
            margin: 15px 0;
            letter-spacing: 0.5px;
          }
          .info-line {
            text-align: center;
            font-size: 11px;
            font-style: italic;
            margin-bottom: 8px;
          }
          .date-info {
            text-align: center;
            font-size: 11px;
            font-style: italic;
            margin-bottom: 5px;
          }
          .date-info.red {
            color: #d32f2f;
          }
          /* 2. BẢNG DỮ LIỆU - Professional Print-Ready */
          table {
            width: 100%;
            border-collapse: collapse;
            font-size: 11px;
            margin-bottom: 20px;
            table-layout: fixed;
          }
          thead th {
            background-color: #f8f9fa;
            font-weight: bold;
            text-transform: uppercase;
            text-align: center;
            padding: 10px 6px;
            border: 1px solid #666666;
            vertical-align: middle;
            font-size: 11px;
          }
          /* Column widths: STT (40px), Mã vạch (100px), Tên SP (250px), Giá (80px), CTKM (60px), Tặng (150px), Giá KM (80px), SL (50px), Thành tiền (90px), Ghi chú (100px) */
          th:nth-child(1), td:nth-child(1) { width: 40px; }  /* STT */
          th:nth-child(2), td:nth-child(2) { width: 100px; } /* Mã vạch */
          th:nth-child(3), td:nth-child(3) { width: 250px; word-wrap: break-word; white-space: normal; } /* Tên SP */
          th:nth-child(4), td:nth-child(4) { width: 80px; } /* Giá */
          th:nth-child(5), td:nth-child(5) { width: 60px; } /* CTKM */
          th:nth-child(6), td:nth-child(6) { width: 150px; word-wrap: break-word; white-space: normal; } /* Tặng Kèm */
          th:nth-child(7), td:nth-child(7) { width: 80px; } /* Giá KM */
          th:nth-child(8), td:nth-child(8) { width: 50px; } /* SL */
          th:nth-child(9), td:nth-child(9) { width: 90px; } /* Thành tiền */
          th:nth-child(10), td:nth-child(10) { width: 100px; word-wrap: break-word; white-space: normal; } /* Ghi chú */
          tbody td {
            border: 1px solid #666666;
            padding: 8px 6px;
            vertical-align: middle;
            font-size: 11px;
            min-height: 25px;
            height: auto;
          }
          tbody tr {
            min-height: 25px;
            height: auto;
          }
          /* Căn lề theo yêu cầu */
          .col-center {
            text-align: center;
          }
          .col-left {
            text-align: left;
            padding-left: 10px;
          }
          .col-right {
            text-align: right;
          }
          tfoot td {
            border: 1px solid #666666;
            padding: 8px 6px;
            font-weight: bold;
            font-size: 11px;
          }
          /* 3. FOOTER & CHỮ KÝ - Professional Print-Ready */
          .footer-date {
            text-align: right;
            font-style: italic;
            margin: 20px 0;
            font-size: 11px;
          }
          .signature-section {
            display: flex;
            justify-content: space-around;
            align-items: flex-start;
            margin-top: 40px; /* Cách bảng 2 dòng trống */
            page-break-inside: avoid;
            width: 100%;
            padding: 0 5%;
          }
          .signature-box {
            flex: 0 0 auto;
            text-align: center;
            min-height: 150px;
            width: 30%;
            max-width: 200px;
          }
          .signature-label {
            font-weight: bold;
            font-size: 12px;
            text-transform: uppercase;
            margin-bottom: 100px; /* Cách tên 5 dòng (khoảng 100px, mỗi dòng 20px) */
            display: block;
            line-height: 1.5;
          }
          .signature-name {
            font-weight: bold;
            font-size: 12px;
            text-transform: uppercase;
            margin-top: 5px;
            line-height: 1.5;
            word-wrap: break-word;
          }
        </style>
      </head>
      <body>
        <!-- 1. THÔNG TIN CÔNG TY: Căn trái, font 10, in thường -->
        <div class="header">
          <div class="company-name">Công ty TNHH Việt Liên</div>
          <div class="company-info">
            Phòng 06, Tầng 19, Tòa nhà Golden King, Số 15 Nguyễn Lương Bằng, P.Tân Mỹ, Q.7, TP.HCM<br>
            ĐT: 028.54136608   FAX: 028.5413.6607<br>
            Email: donhang@vietlien.vn
          </div>
        </div>
        
        <!-- TIÊU ĐỀ: Căn giữa, in hoa, Bold, size 18 -->
        <div class="title">CHƯƠNG TRÌNH KHUYẾN MÃI</div>
        
        <!-- MÃ CTKM VÀ THỜI GIAN: Căn giữa, size 11, in nghiêng -->
        <div class="info-line">
          Hệ thống: <strong>${String(d.system || '').trim()}</strong> - Mã: <strong>${String(d.ctkmName || '').trim()}</strong>
        </div>
        
        <div class="date-info red">
          Thời gian khuyến mãi từ: ${String(d.dateRange || '').trim()}
        </div>
        
        <div class="date-info">
          Thời gian Giao Hàng: ${String(d.delRange || '').trim()}
        </div>
        
        <!-- 2. BẢNG DỮ LIỆU -->
        <table>
          <thead>
            <tr>
              <th style="width:4%">STT</th>
              <th style="width:12%">MÃ VẠCH</th>
              <th style="width:35%">TÊN SẢN PHẨM</th>
              <th style="width:8%">GIÁ (-VAT)</th>
              <th style="width:10%">HÌNH THỨC</th>
              <th style="width:12%">TẶNG KÈM</th>
              <th style="width:8%">GIÁ KM (-VAT)</th>
              <th style="width:6%">SLĐK</th>
              <th style="width:10%">THÀNH TIỀN</th>
              <th style="width:10%">GHI CHÚ</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
          <tfoot>
            <tr>
              <td colspan="7" style="text-align:right;padding:8px 6px;border:1px solid #666666">TỔNG CỘNG (TOTAL):</td>
              <td style="text-align:center;padding:8px 6px;border:1px solid #666666">${typeof d.totalQty === 'string' ? d.totalQty : formatNumber(d.totalQty || 0)}</td>
              <td style="text-align:right;padding:8px 6px;border:1px solid #666666">${typeof d.totalAmount === 'string' ? d.totalAmount : formatNumber(d.totalAmount || 0)}</td>
              <td style="padding:8px 6px;border:1px solid #666666"></td>
            </tr>
          </tfoot>
        </table>
        
        <div class="footer-date">
          TP.HCM, Ngày ${day} tháng ${month} năm ${year}
        </div>
        
        <!-- 3. CHỮ KÝ: Cột B (Người đề nghị), Cột F (Kế toán), Cột I (Giám đốc) -->
        <div class="signature-section">
          <div class="signature-box">
            <span class="signature-label">NGƯỜI ĐỀ NGHỊ</span>
            <div class="signature-name">${String(d.proposer || '').trim().toUpperCase()}</div>
          </div>
          <div class="signature-box">
            <span class="signature-label">KẾ TOÁN</span>
            <div class="signature-name">CHUNG THANH HUỆ</div>
          </div>
          <div class="signature-box">
            <span class="signature-label">GIÁM ĐỐC</span>
            <div class="signature-name">NGUYỄN THỊ THU HẰNG</div>
          </div>
        </div>
      </body>
    </html>`;
}

// --- JSON HELPERS ---
function jsonSuccess(d) {
  return ContentService.createTextOutput(
    JSON.stringify({ success: true, data: d })
  ).setMimeType(ContentService.MimeType.JSON);
}

function jsonError(e) {
  return ContentService.createTextOutput(
    JSON.stringify({ success: false, error: e })
  ).setMimeType(ContentService.MimeType.JSON);
}

// --- DATE HELPERS ---
function parseVNDate(s) {
  if (!s) return null;

  let c = String(s).trim().replace(/-/g, '/');
  const p = c.split('/');

  let d, m, y;
  if (p.length === 3) {
    d = p[0];
    m = p[1] - 1;
    y = p[2];
  } else if (p.length === 2) {
    d = p[0];
    m = p[1] - 1;
    y = new Date().getFullYear();
  } else {
    return null;
  }

  if (isNaN(y) || y < 2000) y = new Date().getFullYear();
  return new Date(y, m, d);
}

function formatTimeRange(s, e) {
  if (!s || !e) return "";
  const f = d => Utilities.formatDate(d, "Asia/Ho_Chi_Minh", "dd/MM/yyyy");
  return `${f(s)} - ${f(e)}`;
}

function parseCustomDateRange(t) {
  if (!t) return { startDate: '', endDate: '' };

  const p = String(t).split('-');
  if (p.length < 2) return { startDate: '', endDate: '' };

  const sRaw = p[0].trim();
  const eRaw = p[1].trim();

  const ep = eRaw.split('/');
  let y = new Date().getFullYear();
  if (ep.length === 3) y = ep[2];

  const iso = (r, yr) => {
    const sp = r.split('/');
    const pd = n => String(n).padStart(2, '0');

    if (sp.length === 3) return `${sp[2]}-${pd(sp[1])}-${pd(sp[0])}`;
    if (sp.length === 2) return `${yr}-${pd(sp[1])}-${pd(sp[0])}`;
    return '';
  };

  let sy = y;
  if (parseInt(sRaw.split('/')[1], 10) > parseInt(eRaw.split('/')[1], 10)) {
    sy = parseInt(y, 10) - 1;
  }

  return {
    startDate: iso(sRaw, sy),
    endDate:   iso(eRaw, y)
  };
}

// --- PRINT TIMELINE (PDF TOÀN BỘ TIMELINE 4 THÁNG) ---
/**
 * Lưu Timeline PDF vào Google Drive
 * In toàn bộ timeline 4 tháng gần nhất theo view mode (system hoặc sku)
 */
function saveTimelinePdf(data) {
  const viewMode = data.viewMode || 'system';
  const months = data.months || [];

  if (!months || !months.length) {
    return { error: "Không có dữ liệu timeline" };
  }

  const baseName = `VietLien_Timeline_${new Date().getTime()}`;
  const files = DriveApp.searchFiles(`title contains '${baseName}' and trashed=false`);

  let count = 0;
  while (files.hasNext()) {
    files.next();
    count++;
  }

  const finalName = `${baseName}_${String(count + 1).padStart(2, '0')}.pdf`;

  const html = createTimelineHtml(data);
  const blob = Utilities.newBlob(html, MimeType.HTML, finalName);
  const file = DriveApp.createFile(blob.getAs(MimeType.PDF)).setName(finalName);

  return { message: finalName, url: file.getUrl() };
}

/**
 * Tạo HTML cho Timeline PDF
 * Hiển thị timeline 4 tháng theo view mode
 */
/**
 * Tạo HTML cho Timeline PDF dạng Gantt table
 * Hiển thị timeline 4 tháng theo dạng bảng (SKU/Brand dọc, tháng ngang)
 */
function createTimelineHtml(payload) {
  const today = new Date();
  const day = String(today.getDate()).padStart(2, '0');
  const monthNum = String(today.getMonth() + 1).padStart(2, '0');
  const year = today.getFullYear();

  // Support legacy call signature: (months, viewMode)
  let months = [];
  let viewMode = 'system';
  let printSystem = '';
  let printBrand = '';

  if (Array.isArray(payload)) {
    months = payload;
    if (arguments.length > 1) viewMode = arguments[1] || 'system';
  } else if (payload && typeof payload === 'object') {
    months = payload.months || [];
    viewMode = payload.viewMode || 'system';
    printSystem = payload.printSystem ? String(payload.printSystem).trim() : '';
    printBrand = payload.printBrand ? String(payload.printBrand).trim() : '';
  }

  // Collect all items and get unique month-year keys
  const allItems = [];
  const monthKeys = new Set();

  months.forEach(monthData => {
    const key = `${monthData.year}-${String(monthData.month).padStart(2, '0')}`;
    monthKeys.add(key);
    (monthData.items || []).forEach(item => {
      allItems.push({ ...item, monthKey: key });
    });
  });

  // Sort months numerically (2025-11 < 2025-12 < 2026-01)
  const sortedMonths = Array.from(monthKeys).sort((a, b) => {
    const [aYear, aMonth] = a.split('-');
    const [bYear, bMonth] = b.split('-');
    const aNum = parseInt(aYear) * 100 + parseInt(aMonth);
    const bNum = parseInt(bYear) * 100 + parseInt(bMonth);
    return aNum - bNum;
  });

  // Build row labels (SKU hoặc Brand) và group items
  let rowData = {};
  
  if (viewMode === 'sku') {
    // Group by SKU
    allItems.forEach(item => {
      const sku = String(item.sku || '').trim();
      if (!rowData[sku]) {
        rowData[sku] = { label: sku, items: {} };
      }
      if (!rowData[sku].items[item.monthKey]) {
        rowData[sku].items[item.monthKey] = [];
      }
      rowData[sku].items[item.monthKey].push(item);
    });
  } else {
    // Group by Brand
    allItems.forEach(item => {
      const brand = String(item.brand || '').trim();
      if (!rowData[brand]) {
        rowData[brand] = { label: brand, items: {} };
      }
      if (!rowData[brand].items[item.monthKey]) {
        rowData[brand].items[item.monthKey] = [];
      }
      rowData[brand].items[item.monthKey].push(item);
    });
  }

  // Sort rows
  const sortedRows = Object.keys(rowData).sort();

  // Helpers to format ISO dates and ranges (define BEFORE pivot code uses them)
  const formatIsoDate = (iso) => {
    if (!iso) return '';
    const parts = String(iso).split('-');
    if (parts.length < 3) return iso;
    return `${String(parts[2]).padStart(2,'0')}/${String(parts[1]).padStart(2,'0')}/${parts[0]}`;
  };

  const formatRange = (sIso, eIso) => {
    if (!sIso && !eIso) return '';
    if (!sIso) return formatIsoDate(eIso);
    if (!eIso) return formatIsoDate(sIso);
    const sp = String(sIso).split('-');
    const ep = String(eIso).split('-');
    if (sp.length >= 3 && ep.length >= 3 && sp[0] === ep[0]) {
      return `${String(sp[2]).padStart(2,'0')}/${String(sp[1]).padStart(2,'0')}-${String(ep[2]).padStart(2,'0')}/${String(ep[1]).padStart(2,'0')}/${ep[0]}`;
    }
    return `${formatIsoDate(sIso)} - ${formatIsoDate(eIso)}`;
  };

  // If printSystem selected, render pivot-style table (columns = CTKM, rows = Brand / SKU)
  if (printSystem) {
    const filtered = allItems.filter(item => {
      const sys = String(item.system || '').trim();
      const br = String(item.brand || '').trim();
      if (!(sys === printSystem || sys.includes(printSystem) || printSystem.includes(sys))) return false;
      if (printBrand && !(br === printBrand || br.includes(printBrand) || printBrand.includes(br))) return false;
      return true;
    });

    const ctkmCols = Array.from(new Set(filtered.map(i => String(i.name || '(Không tên)').trim()))).sort();

    const brandMap = {};
    filtered.forEach(i => {
      const br = String(i.brand || '(No Brand)').trim();
      const sku = String(i.sku || '').trim();
      if (!brandMap[br]) brandMap[br] = new Set();
      if (sku) brandMap[br].add(sku);
    });

    const repForCtkm = name => filtered.find(x => String(x.name || '').trim() === name) || {};

    let html = '';
    html += `
      <html>
        <head>
          <meta charset="UTF-8">
          <style>
            @page { size: A4 landscape; margin: 0.8cm; }
            body { font-family: 'Times New Roman', 'Arial', serif; font-size: 11px; color:#000; margin:0; padding:10px }
            table { width:100%; border-collapse: collapse; }
            th, td { border:1px solid #000; padding:6px; vertical-align:top; }
            th { background:#f0f0f0; font-weight:bold; text-align:center }
            .ctkm-header { font-weight:bold; font-size:12px }
            .small { font-size:11px; color:#333 }
          </style>
        </head>
        <body>
          <h2 style="text-align:center; margin:6px 0 12px 0">BÁO CÁO CTKM - HỆ THỐNG: ${printSystem}</h2>
          <table>
            <thead>
              <tr>
                <th>Hệ thống</th>
                <th>Nhãn</th>
                <th>Sản phẩm (SKU)</th>`;

    ctkmCols.forEach(name => {
      const rep = repForCtkm(name);
      const del = rep.delivery_str || '';
      const startIso = rep.start_date || '';
      const endIso = rep.end_date || '';
      html += `<th><div class="ctkm-header">${name}</div><div class="small">${formatRange(startIso, endIso)}</div><div class="small">${del}</div></th>`;
    });

    html += `</tr></thead><tbody>`;

    Object.keys(brandMap).sort().forEach(brand => {
      const skus = Array.from(brandMap[brand]).sort();
      skus.forEach(sku => {
        html += `<tr><td>${printSystem}</td><td>${brand}</td><td>${sku}</td>`;
        ctkmCols.forEach(name => {
          const found = filtered.find(i => String(i.name || '').trim() === name && String(i.sku || '').trim() === sku);
          if (found) {
            const promo = String(found.promotion || '').trim();
            const gift = String(found.gift || '').trim();
            const val = `${promo || ''}${promo && gift ? ' / ' : ''}${gift ? 'Tặng: ' + gift : ''}`;
            html += `<td style="text-align:center">${val}</td>`;
          } else {
            html += `<td>&nbsp;</td>`;
          }
        });
        html += `</tr>`;
      });
    });

    html += `</tbody></table></body></html>`;
    return html;
  }

  // When exporting all CTKM (no printSystem filter), render month-based pivot table
  // Columns: Month only
  // Rows: System | Brand | SKU
  // Values: CTKM name + delivery date + promo% / gift (in cell)

  // Build month columns structure (just months, no CTKM details in header)
  const monthCols = []; // Array of { monthKey, monthName, ctkms: [ { name, startDate, endDate, deliveryStr } ] }
  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  sortedMonths.forEach(monthKey => {
    const parts = monthKey.split('-');
    const y = parts[0];
    const m = parts[1];
    const monthIdx = parseInt(m);

    if (monthIdx < 1 || monthIdx > 12 || isNaN(monthIdx)) return;

    const monthName = monthNames[monthIdx - 1] + ' ' + y;
    const itemsInMonth = allItems.filter(item => item.monthKey === monthKey);

    // Group by CTKM name in this month (keep all details)
    const ctkmMap = {};
    itemsInMonth.forEach(item => {
      const ctkmName = String(item.name || '(Không tên)').trim();
      if (!ctkmMap[ctkmName]) {
        ctkmMap[ctkmName] = {
          name: ctkmName,
          startDate: item.start_date || '',
          endDate: item.end_date || '',
          deliveryStr: item.delivery_str || '',
          items: []
        };
      }
      ctkmMap[ctkmName].items.push(item);
    });

    monthCols.push({
      monthKey,
      monthName,
      ctkms: Object.values(ctkmMap)
    });
  });

  // Build rows: System | Brand | SKU
  const rowMap = {}; // key: "SYS||BRAND||SKU" -> { system, brand, sku, items: [] }
  allItems.forEach(item => {
    const sys = String(item.system || '').trim();
    const br = String(item.brand || '').trim();
    const sku = String(item.sku || '').trim();
    const key = `${sys}||${br}||${sku}`;

    if (!rowMap[key]) {
      rowMap[key] = { system: sys, brand: br, sku, items: [] };
    }
    rowMap[key].items.push(item);
  });

  const sortedPivotRows = Object.keys(rowMap).sort((a, b) => {
    const aKey = rowMap[a];
    const bKey = rowMap[b];
    if (aKey.system !== bKey.system) return aKey.system.localeCompare(bKey.system);
    if (aKey.brand !== bKey.brand) return aKey.brand.localeCompare(bKey.brand);
    return aKey.sku.localeCompare(bKey.sku);
  });

  // Build header HTML (Month names only)
  let headerHtml = '<tr><th style="border:1px solid #000; padding:8px; background:#e0e0e0; font-weight:bold; text-align:center; min-width:100px">Hệ thống</th>';
  headerHtml += '<th style="border:1px solid #000; padding:8px; background:#e0e0e0; font-weight:bold; text-align:center; min-width:100px">Nhãn</th>';
  headerHtml += '<th style="border:1px solid #000; padding:8px; background:#e0e0e0; font-weight:bold; text-align:center; min-width:100px">SKU</th>';

  monthCols.forEach((col, idx) => {
    const bgColor = idx % 2 === 0 ? '#f0f0f0' : '#e8e8e8';
    headerHtml += `<th style="border:1px solid #000; padding:8px; background:${bgColor}; font-weight:bold; text-align:center; min-width:120px">${col.monthName}</th>`;
  });

  headerHtml += '</tr>';

  // Build body HTML
  let bodyHtml = '';
  sortedPivotRows.forEach(rowKey => {
    const rowInfo = rowMap[rowKey];
    bodyHtml += `<tr><td style="border:1px solid #000; padding:8px; background:#f5f5f5; font-weight:bold; text-align:center; min-width:100px">${rowInfo.system}</td>`;
    bodyHtml += `<td style="border:1px solid #000; padding:8px; background:#f5f5f5; font-weight:bold; text-align:center; min-width:100px">${rowInfo.brand}</td>`;
    bodyHtml += `<td style="border:1px solid #000; padding:8px; background:#f5f5f5; font-weight:bold; text-align:center; min-width:100px">${rowInfo.sku}</td>`;

    monthCols.forEach((col, idx) => {
      const bgColor = idx % 2 === 0 ? '#ffffff' : '#fafafa';
      
      // Find CTKM items for this row + month combination
      const itemsForMonth = rowInfo.items.filter(item => item.monthKey === col.monthKey);
      
      let cellHtml = '<div style="padding:4px; font-size:10px; line-height:1.5">';
      
      if (itemsForMonth.length > 0) {
        // Group by CTKM name
        const ctkmMap = {};
        itemsForMonth.forEach(item => {
          const ctkmName = String(item.name || '(Không tên)').trim();
          if (!ctkmMap[ctkmName]) {
            ctkmMap[ctkmName] = {
              name: ctkmName,
              deliveryStr: item.delivery_str || '',
              items: []
            };
          }
          ctkmMap[ctkmName].items.push(item);
        });

        Object.keys(ctkmMap).forEach(ctkmName => {
          const ctkmInfo = ctkmMap[ctkmName];
          const firstItem = ctkmInfo.items[0] || {};
          const promo = String(firstItem.promotion || '').trim();
          const gift = String(firstItem.gift || '').trim();
          const val = `${promo || ''}${promo && gift ? ' / ' : ''}${gift ? 'Tặng: ' + gift : ''}`;
          
          cellHtml += `<div style="margin-bottom:6px; border-left:2px solid #0066cc; padding-left:4px">
            <div style="font-weight:bold; color:#333">${ctkmName}</div>
            <div style="font-size:9px; color:#666">${ctkmInfo.deliveryStr}</div>
            <div style="font-weight:bold; color:#d32f2f; margin-top:2px">${val}</div>
          </div>`;
        });
      } else {
        cellHtml += '&nbsp;';
      }
      
      cellHtml += '</div>';
      bodyHtml += `<td style="border:1px solid #ccc; padding:6px; text-align:center; background:${bgColor}; vertical-align:top">${cellHtml}</td>`;
    });

    bodyHtml += '</tr>';
  });

  return `
    <html>
      <head>
        <meta charset="UTF-8">
        <style>
          @page {
            size: A4 landscape;
            margin: 0.6cm;
          }
          * {
            box-sizing: border-box;
          }
          body {
            font-family: 'Times New Roman', 'Arial', serif;
            font-size: 10px;
            line-height: 1.3;
            color: #000;
            margin: 0;
            padding: 0;
          }
          .header {
            margin-bottom: 12px;
            text-align: center;
            border-bottom: 2px solid #000;
            padding-bottom: 8px;
          }
          .header h1 {
            margin: 0 0 2px 0;
            font-size: 16px;
            font-weight: bold;
            text-transform: uppercase;
          }
          .header .subtitle {
            margin: 1px 0;
            font-size: 10px;
            font-style: italic;
            color: #666;
          }
          .company-info {
            font-size: 9px;
            margin-top: 4px;
            line-height: 1.3;
          }
          table {
            width: 100%;
            border-collapse: collapse;
            margin: 10px 0;
            font-size: 10px;
          }
          td, th {
            border: 1px solid #000;
            padding: 6px;
            text-align: center;
            vertical-align: middle;
          }
          th {
            background: #e0e0e0;
            font-weight: bold;
          }
          .footer-date {
            text-align: right;
            margin-top: 12px;
            font-style: italic;
            font-size: 9px;
            border-top: 1px solid #ccc;
            padding-top: 6px;
          }
        </style>
      </head>
      <body>
        <!-- HEADER -->
        <div class="header">
          <h1>BẢNG CTKM TOÀN HỆ THỐNG</h1>
          <div class="subtitle">Tất cả Hệ thống - Nhãn - SKU - CTKM</div>
          <div class="company-info">
            Công ty TNHH Việt Liên | ĐT: 028.54136608 | Email: donhang@vietlien.vn
          </div>
        </div>

        <!-- TIMELINE TABLE -->
        <table>
          ${headerHtml}
          ${bodyHtml}
        </table>

        <!-- FOOTER -->
        <div class="footer-date">
          TP.HCM, Ngày ${day} tháng ${monthNum} năm ${year}
        </div>
      </body>
    </html>`;
}

// ============================================
// GUIDELINE MANAGEMENT FUNCTIONS
// ============================================

/**
 * Tải dữ liệu Guideline từ sheet Guideline_CTKM
 * Sheet có cấu trúc: SKU | Tháng | Năm | Ghi chú
 */
function loadGuidelineData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(GUIDELINE_SHEET);
  
  // Tạo sheet mới nếu chưa có
  if (!sheet) {
    sheet = ss.insertSheet(GUIDELINE_SHEET);
    sheet.getRange('A1:D1').setValues([['SKU', 'Tháng', 'Năm', 'Ghi chú']]);
    sheet.getRange('A1:D1').setFontWeight('bold').setBackground('#4CAF50').setFontColor('#ffffff');
    return [];
  }
  
  const values = sheet.getDataRange().getValues();
  const guidelines = [];
  
  for (let i = 1; i < values.length; i++) {
    if (!values[i][0]) continue; // Skip empty rows
    
    guidelines.push({
      sku: String(values[i][0] || '').trim(),
      month: String(values[i][1] || '').trim(),
      year: String(values[i][2] || '').trim(),
      note: String(values[i][3] || '').trim()
    });
  }
  
  return guidelines;
}

/**
 * Lưu Guideline mới hoặc cập nhật guideline hiện có
 * data: { sku, month, year, note }
 */
function saveGuideline(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(GUIDELINE_SHEET);
  
  // Tạo sheet mới nếu chưa có
  if (!sheet) {
    sheet = ss.insertSheet(GUIDELINE_SHEET);
    sheet.getRange('A1:D1').setValues([['SKU', 'Tháng', 'Năm', 'Ghi chú']]);
    sheet.getRange('A1:D1').setFontWeight('bold').setBackground('#4CAF50').setFontColor('#ffffff');
  }
  
  const sku = String(data.sku || '').trim();
  const month = String(data.month || '').trim();
  const year = String(data.year || '').trim();
  const note = String(data.note || '').trim();
  
  if (!sku || !month || !year) {
    return { error: 'Thiếu thông tin SKU, Tháng hoặc Năm!' };
  }
  
  const values = sheet.getDataRange().getValues();
  let rowIdx = -1;
  
  // Tìm dòng có cùng SKU + Tháng + Năm
  for (let i = 1; i < values.length; i++) {
    if (
      cleanStr(values[i][0]) === cleanStr(sku) &&
      String(values[i][1]).trim() === month &&
      String(values[i][2]).trim() === year
    ) {
      rowIdx = i + 1;
      break;
    }
  }
  
  const row = [sku, month, year, note];
  
  if (rowIdx > 0) {
    // Cập nhật dòng hiện có
    sheet.getRange(rowIdx, 1, 1, row.length).setValues([row]);
    return { message: 'Đã cập nhật guideline!' };
  } else {
    // Thêm dòng mới
    sheet.appendRow(row);
    return { message: 'Đã thêm guideline mới!' };
  }
}

/**
 * Xóa Guideline
 * data: { sku, month, year }
 */
function deleteGuideline(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(GUIDELINE_SHEET);
  
  if (!sheet) {
    return { error: 'Không tìm thấy sheet Guideline!' };
  }
  
  const sku = String(data.sku || '').trim();
  const month = String(data.month || '').trim();
  const year = String(data.year || '').trim();
  
  const values = sheet.getDataRange().getValues();
  
  for (let i = 1; i < values.length; i++) {
    if (
      cleanStr(values[i][0]) === cleanStr(sku) &&
      String(values[i][1]).trim() === month &&
      String(values[i][2]).trim() === year
    ) {
      sheet.deleteRow(i + 1);
      return { message: 'Đã xóa guideline!' };
    }
  }
  
  return { error: 'Không tìm thấy guideline để xóa!' };
}

/**
 * Kiểm tra SKU có tuân thủ guideline hay không
 * Trả về: { compliant: true/false, guideline: {...}, actual: {...} }
 */
function checkSkuCompliance(sku, month, year) {
  const guidelines = loadGuidelineData();
  const skuClean = cleanStr(sku);
  
  // Tìm guideline tương ứng
  const guideline = guidelines.find(g => 
    cleanStr(g.sku) === skuClean &&
    String(g.month).trim() === String(month).trim() &&
    String(g.year).trim() === String(year).trim()
  );
  
  if (!guideline) {
    return { compliant: null, guideline: null, actual: null }; // Không có guideline
  }
  
  // Lấy CTKM thực tế từ Data_CTKM
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dS = ss.getSheetByName(DATA_SHEET);
  const dVals = dS.getDataRange().getValues();
  
  let hasPromo = false;
  
  for (let i = 1; i < dVals.length; i++) {
    const row = dVals[i];
    if (cleanStr(row[3]) === skuClean && 
        String(row[0]).includes(`THÁNG ${month}`) &&
        String(row[1]).trim() === year) {
      hasPromo = true;
      break;
    }
  }
  
  return {
    compliant: hasPromo,
    guideline: guideline,
    actual: { sku, month, year, hasPromo }
  };
}