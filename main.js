const { app, BrowserWindow, Menu, dialog, ipcMain } = require('electron');
const ExcelJS = require('exceljs');
const fs = require('node:fs/promises');
const path = require('node:path');

const APP_TITLE = 'PNDA Missions';
let WORKBOOK_PROVINCES = [];

function createMainWindow() {
  const mainWindow = new BrowserWindow({
    title: APP_TITLE,
    icon: path.join(__dirname, 'logo_pnda.ico'),
    width: 1440,
    height: 960,
    minWidth: 1180,
    minHeight: 760,
    backgroundColor: '#f6f1df',
    autoHideMenuBar: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false
    }
  });

  mainWindow.loadFile(path.join(__dirname, 'index.html'));
  buildApplicationMenu(mainWindow);
}

function buildApplicationMenu(mainWindow) {
  const template = [
    {
      label: 'Application',
      submenu: [
        {
          label: 'Formulaire de saisie',
          click: () => mainWindow.loadFile(path.join(__dirname, 'index.html'))
        },
        {
          label: 'Tableau de bord admin',
          click: () => mainWindow.loadFile(path.join(__dirname, 'admin.html'))
        },
        { type: 'separator' },
        { role: 'reload', label: 'Actualiser' },
        { role: 'toggledevtools', label: 'Outils développeur' },
        { type: 'separator' },
        { role: 'quit', label: 'Quitter' }
      ]
    },
    {
      label: 'Édition',
      submenu: [
        { role: 'undo', label: 'Annuler' },
        { role: 'redo', label: 'Rétablir' },
        { type: 'separator' },
        { role: 'cut', label: 'Couper' },
        { role: 'copy', label: 'Copier' },
        { role: 'paste', label: 'Coller' },
        { role: 'selectAll', label: 'Tout sélectionner' }
      ]
    },
    {
      label: 'Fenêtre',
      submenu: [
        { role: 'minimize', label: 'Réduire' },
        { role: 'zoom', label: 'Zoom' },
        { role: 'togglefullscreen', label: 'Plein écran' }
      ]
    }
  ];

  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

ipcMain.handle('dialog:save-file', async (_event, options) => {
  const { canceled, filePath } = await dialog.showSaveDialog({
    title: options?.title || 'Enregistrer le fichier',
    defaultPath: options?.defaultPath,
    filters: options?.filters || []
  });

  if (canceled || !filePath) {
    return { canceled: true };
  }

  await fs.writeFile(filePath, options?.content ?? '', 'utf8');
  return { canceled: false, filePath };
});

ipcMain.handle('excel:export-workbook', async (_event, payload) => {
  const { canceled, filePath } = await dialog.showSaveDialog({
    title: 'Enregistrer le fichier Excel',
    defaultPath: payload?.defaultFileName || `PNDA_dashboard_${new Date().toISOString().slice(0, 10)}.xlsx`,
    filters: [{ name: 'Fichiers Excel', extensions: ['xlsx'] }]
  });

  if (canceled || !filePath) {
    return { canceled: true };
  }

  const workbook = buildDesktopWorkbook(payload || {});
  await workbook.xlsx.writeFile(filePath);
  return { canceled: false, filePath };
});

function buildDesktopWorkbook(payload) {
  const workbook = new ExcelJS.Workbook();
  const exportDate = new Date(payload.exportedAt || Date.now());
  const reports = Array.isArray(payload.reports) ? payload.reports : [];
  const provinces = Array.isArray(payload.provinces) ? payload.provinces : [];
  WORKBOOK_PROVINCES = provinces;
  const signatory = payload.signatory || {};
  const logoPath = path.join(__dirname, 'logo_pnda.png');
  const logoId = workbook.addImage({ filename: logoPath, extension: 'png' });

  workbook.creator = 'PNDA Missions';
  workbook.lastModifiedBy = payload.exportedBy || 'PNDA-SE';
  workbook.created = exportDate;
  workbook.modified = exportDate;
  workbook.company = 'PNDA-SE';
  workbook.subject = 'Suivi des missions';
  workbook.title = 'PNDA Missions';

  const generalSheet = workbook.addWorksheet('General', {
    pageSetup: buildPageSetup(),
    views: [{ state: 'frozen', ySplit: 6 }]
  });
  generalSheet.columns = [{ width: 28 }, { width: 18 }, { width: 16 }, { width: 16 }, { width: 16 }];
  populateSheetHeader(generalSheet, logoId, 'Tableau de bord général', 5);
  fillGeneralWorksheet(generalSheet, reports, provinces, payload.exportedBy, exportDate, signatory);

  provinces.forEach((province) => {
    const sheet = workbook.addWorksheet(sheetNameForProvince(province.label || province.name), {
      pageSetup: buildPageSetup(),
      views: [{ state: 'frozen', ySplit: 6 }]
    });
    sheet.columns = [{ width: 16 }, { width: 16 }, { width: 13 }, { width: 11 }, { width: 34 }, { width: 34 }, { width: 16 }, { width: 16 }];
    populateSheetHeader(sheet, logoId, `Province - ${province.label || province.name}`, 8);
    fillProvinceWorksheet(sheet, reports, province, signatory);
  });

  const missionsSheet = workbook.addWorksheet('Toutes_Missions', {
    pageSetup: buildPageSetup(),
    views: [{ state: 'frozen', ySplit: 6 }]
  });
  missionsSheet.columns = [{ width: 16 }, { width: 22 }, { width: 12 }, { width: 10 }, { width: 12 }, { width: 16 }, { width: 32 }, { width: 32 }, { width: 15 }, { width: 15 }, { width: 15 }];
  populateSheetHeader(missionsSheet, logoId, 'Toutes les missions', 11);
  fillAllMissionsWorksheet(missionsSheet, reports, signatory);

  return workbook;
}

function buildPageSetup() {
  return {
    paperSize: 9,
    orientation: 'landscape',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    margins: {
      left: 0.25,
      right: 0.25,
      top: 0.35,
      bottom: 0.45,
      header: 0.2,
      footer: 0.2
    }
  };
}

function populateSheetHeader(worksheet, logoId, sheetTitle, columnCount) {
  worksheet.properties.defaultRowHeight = 20;
  worksheet.mergeCells(1, 3, 2, columnCount);
  worksheet.mergeCells(3, 3, 3, columnCount);
  worksheet.addImage(logoId, {
    tl: { col: 0.2, row: 0.15 },
    ext: { width: 118, height: 118 }
  });

  worksheet.getRow(1).height = 40;
  worksheet.getRow(2).height = 30;
  worksheet.getRow(3).height = 24;
  worksheet.getRow(4).height = 14;

  const titleCell = worksheet.getCell(1, 3);
  titleCell.value = 'PNDA-SE';
  titleCell.style = styleBrandTitle();

  const subtitleCell = worksheet.getCell(3, 3);
  subtitleCell.value = `Programme National de Développement Agricole - ${sheetTitle}`;
  subtitleCell.style = styleBrandSubtitle();
}

function fillGeneralWorksheet(worksheet, reports, provinces, exportedBy, exportDate, signatory) {
  const missions = reports.flatMap((report) => report.missions || []);
  let row = 5;

  addMergedTitle(worksheet, row++, 5, 'PNDA-SE - Tableau de bord général');
  addKeyValueRow(worksheet, row++, 'Date export', formatDateFr(exportDate), 2);
  addKeyValueRow(worksheet, row++, 'Utilisateur', exportedBy || '—', 2);
  row++;

  addHeaderRow(worksheet, row++, ['Indicateur', 'Valeur']);
  addDataRow(worksheet, row++, ['Rapports consolidés', reports.length], ['text', 'integer']);
  addDataRow(worksheet, row++, ['Missions consolidées', missions.length], ['text', 'integer']);
  addDataRow(worksheet, row++, ['Montant total USD', missions.reduce((sum, mission) => sum + Number(mission.montant_usd || 0), 0)], ['text', 'currency']);
  addDataRow(worksheet, row++, ['Avance totale USD', missions.reduce((sum, mission) => sum + Number(mission.avance_usd || 0), 0)], ['text', 'currency']);
  addDataRow(worksheet, row++, ['Solde total USD', missions.reduce((sum, mission) => sum + Number(mission.solde_usd || 0), 0)], ['text', 'currency']);
  row++;

  addMergedTitle(worksheet, row++, 5, 'Consolidation par province');
  addHeaderRow(worksheet, row++, ['Province', 'Rapports', 'Missions', 'Montant USD', 'Solde USD']);
  provinces.forEach((province) => {
    const summary = summarizeProvince(reports, provinces, province.name);
    addDataRow(worksheet, row++, [summary.label, summary.reports, summary.missions, summary.montant, summary.solde], ['text', 'integer', 'integer', 'currency', 'currency']);
  });
  row++;

  addMergedTitle(worksheet, row++, 5, 'Derniers rapports');
  addHeaderRow(worksheet, row++, ['Date', 'Province', 'Rapporteur', 'Missions', 'Source']);
  reports.slice(0, 20).forEach((report) => {
    addDataRow(worksheet, row++, [formatDateFr(report.savedAt), report.meta?.province || '—', report.meta?.rapporteur || '—', report.meta?.total_missions || report.missions?.length || 0, report.source || 'stockage local'], ['dateText', 'text', 'text', 'integer', 'text']);
  });

  appendSignatureBlock(worksheet, row + 2, 5, signatory);
}

function fillProvinceWorksheet(worksheet, reports, province, signatory) {
  const provinceReports = reports.filter((report) => getReportProvinceName(report, province.name) === province.name);
  const missions = provinceReports.flatMap((report) => (report.missions || []).map((mission) => ({ ...mission, __report: report })));
  let row = 5;

  addMergedTitle(worksheet, row++, 8, `Province - ${province.label || province.name}`);
  addKeyValueRow(worksheet, row++, 'Rapports', provinceReports.length, 2, 'integer');
  addKeyValueRow(worksheet, row++, 'Missions', missions.length, 2, 'integer');
  row++;

  addMergedTitle(worksheet, row++, 6, 'Rapports');
  addHeaderRow(worksheet, row++, ['Date', 'Rapporteur', 'Trimestre', 'Année', 'Montant USD', 'Solde USD']);
  provinceReports.forEach((report) => {
    addDataRow(worksheet, row++, [formatDateFr(report.savedAt), report.meta?.rapporteur || '—', report.meta?.trimestre || '—', report.meta?.annee || '—', Number(report.meta?.total_montant_usd || 0), Number(report.meta?.total_solde_usd || 0)], ['dateText', 'text', 'text', 'integer', 'currency', 'currency']);
  });
  row++;

  addMergedTitle(worksheet, row++, 8, 'Détail des missions');
  addHeaderRow(worksheet, row++, ['Date mission', 'Numero', 'Nature', 'Objectif', 'Rapporteur', 'Montant USD', 'Avance USD', 'Solde USD']);
  missions.forEach((mission) => {
    addDataRow(worksheet, row++, [formatDateFr(mission.date || mission.__report.savedAt), mission.numero || '—', mission.nature || '—', mission.objectif || '—', mission.__report.meta?.rapporteur || '—', Number(mission.montant_usd || 0), Number(mission.avance_usd || 0), Number(mission.solde_usd || 0)], ['dateText', 'text', 'text', 'text', 'text', 'currency', 'currency', 'currency']);
  });

  appendSignatureBlock(worksheet, row + 2, 8, signatory);
}

function fillAllMissionsWorksheet(worksheet, reports, signatory) {
  const missions = reports.flatMap((report) => (report.missions || []).map((mission) => ({ ...mission, __report: report })));
  let row = 5;

  addMergedTitle(worksheet, row++, 11, 'Toutes les missions');
  addHeaderRow(worksheet, row++, ['Province', 'Rapporteur', 'Trimestre', 'Année', 'Numero', 'Mission ref', 'Nature', 'Objectif', 'Montant USD', 'Avance USD', 'Solde USD']);
  missions.forEach((mission) => {
    addDataRow(worksheet, row++, [getReportProvinceLabel(mission.__report) || '—', mission.__report.meta?.rapporteur || '—', mission.__report.meta?.trimestre || '—', mission.__report.meta?.annee || '—', mission.numero || '—', mission.mission_reference_id || '—', mission.nature || '—', mission.objectif || '—', Number(mission.montant_usd || 0), Number(mission.avance_usd || 0), Number(mission.solde_usd || 0)], ['text', 'text', 'text', 'integer', 'text', 'text', 'text', 'text', 'currency', 'currency', 'currency']);
  });

  appendSignatureBlock(worksheet, row + 2, 11, signatory);
}

function appendSignatureBlock(worksheet, startRow, totalColumns, signatory) {
  const name = signatory?.name || '................................';
  const position = signatory?.position || '................................';
  const location = signatory?.location || '........................';
  const exportDate = formatDateFr(new Date());
  const startColumn = Math.max(totalColumns - 1, 1);

  worksheet.mergeCells(startRow, startColumn, startRow, totalColumns);
  worksheet.mergeCells(startRow + 1, startColumn, startRow + 1, totalColumns);
  worksheet.mergeCells(startRow + 2, startColumn, startRow + 2, totalColumns);
  worksheet.mergeCells(startRow + 4, startColumn, startRow + 4, totalColumns);
  worksheet.mergeCells(startRow + 5, startColumn, startRow + 5, totalColumns);

  worksheet.getCell(startRow, startColumn).value = `Fait à ${location}, le ${exportDate}`;
  worksheet.getCell(startRow, startColumn).style = styleSignatureValue();
  worksheet.getCell(startRow + 1, startColumn).value = 'Signature';
  worksheet.getCell(startRow + 1, startColumn).style = styleSignatureLabel();
  worksheet.getCell(startRow + 2, startColumn).value = '_______________________________';
  worksheet.getCell(startRow + 2, startColumn).style = styleSignatureLine();
  worksheet.getCell(startRow + 4, startColumn).value = `Nom: ${name}`;
  worksheet.getCell(startRow + 4, startColumn).style = styleSignatureValue();
  worksheet.getCell(startRow + 5, startColumn).value = `Poste: ${position}`;
  worksheet.getCell(startRow + 5, startColumn).style = styleSignatureValue();
}

function addMergedTitle(worksheet, rowNumber, totalColumns, value) {
  worksheet.mergeCells(rowNumber, 1, rowNumber, totalColumns);
  const cell = worksheet.getCell(rowNumber, 1);
  cell.value = value;
  cell.style = styleSectionTitle();
}

function addKeyValueRow(worksheet, rowNumber, label, value, valueColumn = 2, valueStyle = 'text') {
  const labelCell = worksheet.getCell(rowNumber, 1);
  labelCell.value = label;
  labelCell.style = styleSubtitle();

  const valueCell = worksheet.getCell(rowNumber, valueColumn);
  valueCell.value = value;
  valueCell.style = resolveStyle(valueStyle);
}

function addHeaderRow(worksheet, rowNumber, headers) {
  headers.forEach((header, index) => {
    const cell = worksheet.getCell(rowNumber, index + 1);
    cell.value = header;
    cell.style = styleTableHeader();
  });
}

function addDataRow(worksheet, rowNumber, values, styles) {
  values.forEach((value, index) => {
    const cell = worksheet.getCell(rowNumber, index + 1);
    cell.value = value;
    cell.style = resolveStyle(styles[index] || 'text');
  });
}

function summarizeProvince(reports, provinces, provinceName) {
  const provinceReports = reports.filter((report) => getReportProvinceName(report, provinceName) === provinceName);
  const missions = provinceReports.flatMap((report) => report.missions || []);
  return {
    label: provinces.find((province) => province.name === provinceName)?.label || provinceName,
    reports: provinceReports.length,
    missions: missions.length,
    montant: missions.reduce((sum, mission) => sum + Number(mission.montant_usd || 0), 0),
    solde: missions.reduce((sum, mission) => sum + Number(mission.solde_usd || 0), 0)
  };
}

function getReportProvinceName(report, fallback = '') {
  if (report?.meta?.province) {
    return report.meta.province;
  }

  const provinceLabel = report?.meta?.province_label;
  return WORKBOOK_PROVINCES.find((province) => province.label === provinceLabel)?.name || provinceLabel || fallback;
}

function getReportProvinceLabel(report) {
  const provinceName = getReportProvinceName(report);
  return WORKBOOK_PROVINCES.find((province) => province.name === provinceName)?.label || report?.meta?.province_label || provinceName || '—';
}

function sheetNameForProvince(label) {
  return String(label || 'Province').replace(/[\\/*?:\[\]]/g, '_').slice(0, 31);
}

function formatDateFr(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return '—';
  }

  return date.toLocaleDateString('fr-FR');
}

function resolveStyle(name) {
  const map = {
    text: styleText(),
    subtitle: styleSubtitle(),
    integer: styleInteger(),
    currency: styleCurrency(),
    dateText: styleDateText()
  };
  return map[name] || styleText();
}

function styleBrandTitle() {
  return {
    font: { name: 'Calibri', size: 16, bold: true, color: { argb: 'FF123524' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCEEE2' } },
    alignment: { vertical: 'middle', horizontal: 'left' },
    border: { bottom: { style: 'thin', color: { argb: 'FF8FB89F' } } }
  };
}

function styleBrandSubtitle() {
  return {
    font: { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF2F6A48' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEDF7F0' } },
    alignment: { vertical: 'middle', horizontal: 'left', wrapText: true },
    border: { bottom: { style: 'thin', color: { argb: 'FFB9D7C3' } } }
  };
}

function styleSectionTitle() {
  return {
    font: { name: 'Calibri', size: 14, bold: true, color: { argb: 'FF123524' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCFE5D6' } },
    alignment: { vertical: 'middle', horizontal: 'left', wrapText: true },
    border: { bottom: { style: 'thin', color: { argb: 'FF8FB89F' } } }
  };
}

function styleSubtitle() {
  return {
    font: { name: 'Calibri', size: 12, bold: true, color: { argb: 'FF123524' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEEF7F1' } },
    alignment: { vertical: 'middle', horizontal: 'left' }
  };
}

function styleTableHeader() {
  return {
    font: { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF123524' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB8D7C1' } },
    alignment: { vertical: 'middle', horizontal: 'center', wrapText: true },
    border: { bottom: { style: 'thin', color: { argb: 'FF7DAA8C' } } }
  };
}

function styleText() {
  return {
    font: { name: 'Calibri', size: 11, color: { argb: 'FF1B2019' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFAFDFC' } },
    alignment: { vertical: 'top', horizontal: 'left', wrapText: true }
  };
}

function styleInteger() {
  return {
    font: { name: 'Calibri', size: 11, color: { argb: 'FF1B2019' } },
    alignment: { vertical: 'middle', horizontal: 'right' },
    numFmt: '0'
  };
}

function styleCurrency() {
  return {
    font: { name: 'Calibri', size: 11, color: { argb: 'FF1B2019' } },
    alignment: { vertical: 'middle', horizontal: 'right' },
    numFmt: '#,##0.00'
  };
}

function styleDateText() {
  return {
    font: { name: 'Calibri', size: 11, color: { argb: 'FF1B2019' } },
    alignment: { vertical: 'middle', horizontal: 'center' }
  };
}

function styleSignatureLabel() {
  return {
    font: { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF123524' } },
    alignment: { vertical: 'middle', horizontal: 'left' }
  };
}

function styleSignatureLine() {
  return {
    font: { name: 'Calibri', size: 11, color: { argb: 'FF123524' } },
    alignment: { vertical: 'middle', horizontal: 'left' }
  };
}

function styleSignatureValue() {
  return {
    font: { name: 'Calibri', size: 11, color: { argb: 'FF123524' } },
    alignment: { vertical: 'middle', horizontal: 'left', wrapText: true }
  };
}

app.whenReady().then(() => {
  createMainWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createMainWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});