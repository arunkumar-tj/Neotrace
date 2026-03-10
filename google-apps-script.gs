/**
 * Deploy as Web App:
 * 1) Extensions -> Apps Script (from the target spreadsheet)
 * 2) Paste this file, Save
 * 3) Deploy -> New deployment -> Web app
 * 4) Execute as: Me, Who has access: Anyone
 * 5) Copy Web App URL and paste it in app Admin -> Google Sheet Sync -> Save URL
 */

var TARGET_SHEET_ID = '1RAzOiM_NX9wYqrkRdVBo0Ohm5sKkMGtnTE8ibwjcc3E';
var TARGET_GID = 212071646;
var DEBUG_SHEET_NAME = 'SyncDebug';
var QC_SHEET_NAME = 'QCReports';
// Optional: set a specific Drive folder ID for QC PDFs. Leave empty to use My Drive root.
var QC_DRIVE_FOLDER_ID = '';
var HEADERS = [
  'id', 'action', 'type', 'orderNo', 'frameNo', 'inspector', 'timestamp',
  'batteryNo', 'chargerNo', 'motorNo', 'createdAt', 'updatedAt', 'syncedAt'
];

function doPost(e) {
  try {
    var payload = parsePayload_(e);
    var action = String(payload.action || '').toLowerCase();
    var record = payload.record || {};

    logSyncDebug_('REQUEST', JSON.stringify({ action: action, id: record.id || '', orderNo: record.orderNo || '' }));

    var sheet = getTargetSheet_();
    ensureHeaders_(sheet);

    if (action === 'qc_submit') {
      var qcResult = handleQcSubmit_(payload);
      logSyncDebug_('QC_PDF', qcResult.fileId || '');
      return ok_('qc_saved');
    }

    if (!record.id) {
      logSyncDebug_('SKIP', 'missing id');
      return ok_('missing id');
    }

    if (action === 'delete') {
      deleteById_(sheet, record.id);
      logSyncDebug_('DELETE', record.id);
      return ok_('deleted');
    }

    upsertRecord_(sheet, action || 'upsert', record, payload.syncedAt || new Date().toISOString());
    logSyncDebug_('UPSERT', record.id);
    return ok_('upserted');
  } catch (err) {
    logSyncDebug_('ERROR', String(err));
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function parsePayload_(e) {
  if (e && e.parameter && e.parameter.payload) {
    return JSON.parse(e.parameter.payload);
  }

  var body = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
  return JSON.parse(body);
}

function getTargetSheet_() {
  var ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (String(sheets[i].getSheetId()) === String(TARGET_GID)) return sheets[i];
  }

  // Fallback to first sheet to avoid hard failures when gid changes.
  if (sheets.length > 0) {
    logSyncDebug_('WARN', 'Target gid not found, using first sheet: ' + sheets[0].getName());
    return sheets[0];
  }

  throw new Error('No sheets found in spreadsheet: ' + TARGET_SHEET_ID);
}

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    return;
  }

  var existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var needsReset = false;
  if (existing.length < HEADERS.length) {
    needsReset = true;
  } else {
    for (var i = 0; i < HEADERS.length; i++) {
      if (String(existing[i] || '') !== HEADERS[i]) {
        needsReset = true;
        break;
      }
    }
  }

  if (needsReset) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
}

function findRowById_(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function deleteById_(sheet, id) {
  var row = findRowById_(sheet, id);
  if (row > 1) sheet.deleteRow(row);
}

function upsertRecord_(sheet, action, record, syncedAt) {
  var row = [
    record.id || '',
    action || 'upsert',
    record.type || '',
    record.orderNo || '',
    record.frameNo || '',
    record.inspector || '',
    record.timestamp || '',
    record.batteryNo || '',
    record.chargerNo || '',
    record.motorNo || '',
    record.createdAt || '',
    record.updatedAt || '',
    syncedAt || ''
  ];

  var existingRow = findRowById_(sheet, record.id);
  if (existingRow > 1) {
    sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
}

function logSyncDebug_(status, message) {
  try {
    var ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    var sh = ss.getSheetByName(DEBUG_SHEET_NAME);
    if (!sh) sh = ss.insertSheet(DEBUG_SHEET_NAME);
    if (sh.getLastRow() === 0) sh.appendRow(['timestamp', 'status', 'message']);
    sh.appendRow([new Date().toISOString(), status, message]);
  } catch (e) {
    // Swallow logging failures to avoid blocking sync.
  }
}

function ok_(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, message: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleQcSubmit_(payload) {
  var ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
  var sh = ss.getSheetByName(QC_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(QC_SHEET_NAME);

  var qcHeaders = ['id', 'template', 'docId', 'orderNo', 'customerName', 'frameNo', 'inspector',
                    'qcDate', 'qcStartTime', 'qcEndTime', 'headerDataJson', 'sectionsJson',
                    'notes', 'submittedAt', 'pdfFileId', 'pdfUrl'];
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, qcHeaders.length).setValues([qcHeaders]);
  }

  var sections = payload.sections || [];
  var headerData = payload.headerData || {};
  var fileName = 'QC_' + (payload.template || 'UNKNOWN') + '_' + (payload.orderNo || 'NO_ORDER') + '_' + (payload.qcDate || new Date().toISOString().slice(0, 10));

  // --- Build professional PDF via Google Doc ---
  var doc = DocumentApp.create(fileName);
  var body = doc.getBody();

  // Set page margins
  body.setMarginTop(36);
  body.setMarginBottom(36);
  body.setMarginLeft(40);
  body.setMarginRight(40);

  // Title
  var title = body.appendParagraph(payload.templateTitle || 'Neotrace QC Report');
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  // Doc ID
  var docIdPara = body.appendParagraph('Document ID: ' + (payload.docId || '-'));
  docIdPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  docIdPara.editAsText().setFontSize(9).setItalic(true);

  body.appendParagraph('');

  // Header info table
  var hdrRows = [
    ['Order No:', payload.orderNo || '-', 'Customer Name:', payload.customerName || '-'],
    ['Date:', payload.qcDate || '-', 'QC Person:', payload.inspector || '-'],
    ['Frame/Chassis No:', payload.frameNo || '-', 'QC Start Time:', payload.qcStartTime || '-'],
    ['', '', 'QC End Time:', payload.qcEndTime || '-']
  ];

  // Add template-specific header data
  var hdrKeys = Object.keys(headerData);
  for (var h = 0; h < hdrKeys.length; h += 2) {
    var row = [hdrKeys[h] + ':', headerData[hdrKeys[h]] || '-'];
    if (h + 1 < hdrKeys.length) {
      row.push(hdrKeys[h + 1] + ':');
      row.push(headerData[hdrKeys[h + 1]] || '-');
    } else {
      row.push('');
      row.push('');
    }
    hdrRows.push(row);
  }

  var hdrTable = body.appendTable(hdrRows);
  styleHeaderTable_(hdrTable);

  body.appendParagraph('');

  // Render each section
  sections.forEach(function (section) {
    var secTitle = body.appendParagraph(section.name);
    secTitle.setHeading(DocumentApp.ParagraphHeading.HEADING2);

    if (section.packing) {
      // Packing checklist table
      var packRows = [['S.No', 'Box', 'Item', 'Checked', 'Remark']];
      section.items.forEach(function (item) {
        packRows.push([
          String(item.no || ''),
          item.box || '',
          item.param || '',
          item.checked ? 'YES' : (item.value || '-'),
          item.remark || ''
        ]);
      });
      var packTable = body.appendTable(packRows);
      styleChecklistTable_(packTable);
    } else {
      // QC checklist table
      var rows = [['S.No', 'Checking Parameter', 'Value / Status', 'Remark']];
      section.items.forEach(function (item) {
        var displayValue = item.value || '-';
        if (item.type === 'ok' && item.checked) {
          displayValue = 'OK (' + (item.standard || '') + ')';
        } else if (item.type === 'y' && item.checked) {
          displayValue = 'Yes';
        } else if (item.type === 'yn') {
          displayValue = item.value || '-';
        }
        rows.push([
          String(item.no || ''),
          item.param || '',
          displayValue,
          item.remark || ''
        ]);
      });
      var table = body.appendTable(rows);
      styleChecklistTable_(table);
    }

    body.appendParagraph('');
  });

  // Notes
  if (payload.notes) {
    body.appendParagraph('Overall Notes').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(payload.notes);
    body.appendParagraph('');
  }

  // Footer
  var footer = body.appendParagraph('Submitted: ' + (payload.submittedAt || new Date().toISOString()) + '  |  Generated by Neotrace');
  footer.editAsText().setFontSize(8).setItalic(true).setForegroundColor('#888888');
  footer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  doc.saveAndClose();

  // Convert to PDF
  var docFile = DriveApp.getFileById(doc.getId());
  var pdfBlob = docFile.getAs(MimeType.PDF).setName(fileName + '.pdf');

  var folder = null;
  if (QC_DRIVE_FOLDER_ID) {
    folder = DriveApp.getFolderById(QC_DRIVE_FOLDER_ID);
  } else {
    folder = DriveApp.getRootFolder();
  }
  var pdfFile = folder.createFile(pdfBlob);

  // Make PDF non-editable: private view-only
  pdfFile.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW);

  // Remove editable source doc, keep only PDF.
  docFile.setTrashed(true);

  // Log to QC sheet
  sh.appendRow([
    payload.id || '',
    payload.template || '',
    payload.docId || '',
    payload.orderNo || '',
    payload.customerName || '',
    payload.frameNo || '',
    payload.inspector || '',
    payload.qcDate || '',
    payload.qcStartTime || '',
    payload.qcEndTime || '',
    JSON.stringify(headerData),
    JSON.stringify(sections),
    payload.notes || '',
    payload.submittedAt || new Date().toISOString(),
    pdfFile.getId(),
    pdfFile.getUrl()
  ]);

  return { fileId: pdfFile.getId(), fileUrl: pdfFile.getUrl() };
}

function styleHeaderTable_(table) {
  for (var r = 0; r < table.getNumRows(); r++) {
    var row = table.getRow(r);
    for (var c = 0; c < row.getNumCells(); c++) {
      var cell = row.getCell(c);
      cell.editAsText().setFontSize(9);
      if (c % 2 === 0) {
        cell.editAsText().setBold(true);
        cell.setBackgroundColor('#f0f0f0');
      }
      cell.setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(6).setPaddingRight(6);
    }
  }
}

function styleChecklistTable_(table) {
  // Style header row
  if (table.getNumRows() > 0) {
    var hdrRow = table.getRow(0);
    for (var c = 0; c < hdrRow.getNumCells(); c++) {
      var cell = hdrRow.getCell(c);
      cell.setBackgroundColor('#111111');
      cell.editAsText().setForegroundColor('#ffffff').setBold(true).setFontSize(8);
      cell.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(6).setPaddingRight(6);
    }
  }
  // Style data rows
  for (var r = 1; r < table.getNumRows(); r++) {
    var row = table.getRow(r);
    var bgColor = r % 2 === 0 ? '#f9f9f9' : '#ffffff';
    for (var c = 0; c < row.getNumCells(); c++) {
      var cell = row.getCell(c);
      cell.setBackgroundColor(bgColor);
      cell.editAsText().setFontSize(8);
      cell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(6).setPaddingRight(6);
    }
  }
}
