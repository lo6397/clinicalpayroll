const FORCE_REPLACE_CURRENT_PAYDATE_FILES = true;
const ENABLE_BIWEEKLY_CHECK = true;

function runAllTimecards() {
  if (ENABLE_BIWEEKLY_CHECK && !isScheduledRunDay()) {
    Logger.log('Not the scheduled biweekly Wednesday. Skipping.');
    return;
  }

  createInfusionTimecards();
  createPVTimecards();
  // createCorporateTimecards() is not implemented yet — create Corporate timecards manually until it exists.
  // createCorporateTimecards();
}

/************ INFUSION ************/

function createInfusionTimecards() {
  runTimecardWorkflow(
    '1SsppUL9g2QngpoSDFVrpazreojCPmN8uZSm9hCBYZrI',
    [
      '1abALtYC_Bcdnl-J06wtOxyJQI6XFdpQs',
      '1qF4Nv_MLLOEbd4i8nb6PF_2TsnryH0o2',
      '1B4ZbEOPkQ8GW-RETHpWLcQDob137SRhG',
      '1Y3kbrNn_v2E8oyJQYCiH0Swj8m2bmwna'
    ],
    '11e24KTEMKtPKEqDKY8gdwp5hsB6H1Wp1',
    '_Time_Cards',
    {
      itemColumn: 2,
      beginningColumn: 3,
      endingColumn: 5,
      firstItemRow: 3,
      bannerRange: 'C1:F1',
      payPeriodRange: 'C1:E1'
    }
  );
}

/************ VASCULAR ************/

function createPVTimecards() {
  runTimecardWorkflow(
    '1oE2Qjxn0NqhlMDCVt_Rk85ZFAlu7KDwPhCbtGQNp7so',
    ['185swROseZQ4U1nRhHtD-pFNIRx0DIz19'],
    '1kF1oZsaaXQGUNArRM58xp1Ac7tqLIDzG',
    '_PV',
    {
      itemColumn: 1,
      beginningColumn: 2,
      endingColumn: 4,
      firstItemRow: 3,
      bannerRange: 'C1:F1',
      payPeriodRange: 'C1:E1'
    }
  );
}

/************ MAIN WORKFLOW ************/

function runTimecardWorkflow(templateId, parentFolderIds, adminFolderId, archiveSuffix, inventoryConfig) {
  const timecardSheetName = 'TimeCardNotes';
  const inventorySheetName = 'Inventory';

  const runInfo = getCurrentRunInfo();

  const templateFile = DriveApp.getFileById(templateId);
  const adminFolder = DriveApp.getFolderById(adminFolderId);

  const previousArchiveFolder = getOrCreateFolder(
    adminFolder,
    formatPeriodFolderName(runInfo.previousPeriodStart, runInfo.previousPeriodEnd, archiveSuffix)
  );

  getOrCreateFolder(
    adminFolder,
    formatPeriodFolderName(runInfo.currentPeriodStart, runInfo.currentPeriodEnd, archiveSuffix)
  );

  parentFolderIds.forEach(function(parentFolderId) {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const employeeFolders = parentFolder.getFolders();

    while (employeeFolders.hasNext()) {
      const employeeFolder = employeeFolders.next();
      const employeeName = employeeFolder.getName();

      const priorTimecard = getMostRecentTimecard(employeeFolder);

      const fileName =
        employeeName +
        ' - Timecard - Pay Date ' +
        formatDateForFile(runInfo.payDate);

      const existing = employeeFolder.getFilesByName(fileName);

      if (existing.hasNext() && !FORCE_REPLACE_CURRENT_PAYDATE_FILES) {
        Logger.log('Already exists, skipped: ' + fileName);
        continue;
      }

      while (existing.hasNext()) {
        const oldCurrentFile = existing.next();
        oldCurrentFile.moveTo(previousArchiveFolder);
      }

      try {
        const newFile = templateFile.makeCopy(fileName, employeeFolder);
        applyFolderEditorSharingToFile_(newFile, employeeFolder);

        setupNewTimecard(
          newFile,
          timecardSheetName,
          inventorySheetName,
          runInfo.currentPeriodStart,
          runInfo.currentPeriodEnd,
          priorTimecard,
          inventoryConfig
        );

        archiveExistingTimecards(employeeFolder, previousArchiveFolder, newFile.getId());

        Logger.log('Created: ' + fileName);
      } catch (err) {
        Logger.log('ERROR creating ' + fileName + ': ' + err);
      }
    }
  });
}

/************ SETUP ************/

function setupNewTimecard(
  newFile,
  timecardSheetName,
  inventorySheetName,
  currentPeriodStart,
  currentPeriodEnd,
  priorTimecard,
  inventoryConfig
) {
  const spreadsheet = SpreadsheetApp.openById(newFile.getId());
fixTimecardProtections(spreadsheet);
  // Pay period banner on TimeCardNotes sheet
  const timecardSheet = spreadsheet.getSheetByName(timecardSheetName);
  if (timecardSheet) {
    if (inventoryConfig && inventoryConfig.payPeriodRange) {
      try {
        const payPeriodRange = timecardSheet.getRange(inventoryConfig.payPeriodRange);

        payPeriodRange.setValue(
          'Pay Period: ' +
          formatPeriodText(currentPeriodStart) +
          ' - ' +
          formatPeriodText(currentPeriodEnd)
        );

        payPeriodRange.setFontWeight('bold');
        payPeriodRange.setFontSize(14);
        payPeriodRange.setBackground('#FFF2CC');
        payPeriodRange.setHorizontalAlignment('center');
        payPeriodRange.setVerticalAlignment('middle');
      } catch (err) {
        Logger.log('WARNING: Could not set pay period banner on ' + newFile.getName() + ': ' + err);
      }
    }
  }

  // Inventory tab: wipe and replace with the ClickUp inventory form link
  const inventorySheet = spreadsheet.getSheetByName(inventorySheetName);
  if (inventorySheet) {
    try {
      const FORM_URL = 'https://forms.clickup.com/9017962545/f/8cr6c1h-8417/S1B47HI0Z0GR2VM7L2';
      const runInfo = getCurrentRunInfo();
      const payPeriodDates =
        formatPeriodText(runInfo.currentPeriodStart) + ' - ' + formatPeriodText(runInfo.currentPeriodEnd);

      inventorySheet
        .getRange(1, 1, inventorySheet.getMaxRows(), inventorySheet.getMaxColumns())
        .breakApart();
      inventorySheet.clear();

      inventorySheet.getRange('A1:E1').merge();
      const headerCell = inventorySheet.getRange('A1');
      headerCell.setValue('Inventory — Pay Period: ' + payPeriodDates);
      headerCell.setFontWeight('bold');
      headerCell.setFontSize(14);
      headerCell.setBackground('#FFF2CC');
      headerCell.setHorizontalAlignment('center');

      inventorySheet.getRange('A3:E3').merge();
      const linkCell = inventorySheet.getRange('A3');
      linkCell.setFormula('=HYPERLINK("' + FORM_URL + '", "ENTER YOUR INVENTORY HERE")');
      linkCell.setFontWeight('bold');
      linkCell.setFontSize(18);
      linkCell.setHorizontalAlignment('center');
      inventorySheet.setRowHeight(3, 50);

      for (let c = 1; c <= 5; c++) {
        inventorySheet.autoResizeColumn(c);
      }
    } catch (err) {
      Logger.log('WARNING: Could not rebuild inventory tab on ' + newFile.getName() + ': ' + err);
    }
  }
}

/************ INVENTORY ************/

function copyPreviousInventoryCounts(priorFile, newSpreadsheet, inventorySheetName, config) {
  const oldSpreadsheet = SpreadsheetApp.openById(priorFile.getId());

  const oldInventory = oldSpreadsheet.getSheetByName(inventorySheetName);
  const newInventory = newSpreadsheet.getSheetByName(inventorySheetName);

  if (!oldInventory || !newInventory) return;

  const itemValues = oldInventory
    .getRange(
      config.firstItemRow,
      config.itemColumn,
      oldInventory.getMaxRows() - config.firstItemRow + 1,
      1
    )
    .getValues();

  const actualItemCount = itemValues.filter(function(row) {
    return row[0] !== '';
  }).length;

  if (actualItemCount < 1) return;

  const oldEndingCounts = oldInventory
    .getRange(config.firstItemRow, config.endingColumn, actualItemCount, 1)
    .getValues();

  const oldBeginningCounts = oldInventory
    .getRange(config.firstItemRow, config.beginningColumn, actualItemCount, 1)
    .getValues();

  const employee = priorFile.getName().split(' - Timecard')[0].trim();

  const newBeginningCounts = [];
  for (let i = 0; i < actualItemCount; i++) {
    const itemName = itemValues[i][0];
    const ending = oldEndingCounts[i][0];
    const beginning = oldBeginningCounts[i][0];

    if (ending !== '' && ending !== null) {
      newBeginningCounts.push([ending]);
    } else if (beginning !== '' && beginning !== null) {
      newBeginningCounts.push([beginning]);
      Logger.log(employee + ' - ' + itemName + ' from Beginning — flag review');
    } else {
      newBeginningCounts.push(['']);
    }
  }

  newInventory
    .getRange(config.firstItemRow, config.beginningColumn, actualItemCount, 1)
    .setValues(newBeginningCounts);

  newInventory
    .getRange(config.firstItemRow, config.beginningColumn, actualItemCount, 1)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  newInventory
    .getRange(config.firstItemRow, config.endingColumn, actualItemCount, 1)
    .clearContent();
}


/************ ARCHIVE ************/

function archiveExistingTimecards(employeeFolder, archiveFolder, newFileId) {
  const files = employeeFolder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();

    const shouldArchive =
      file.getId() !== newFileId &&
      (
        fileName.includes('Timecard') ||
        fileName.includes('Summary of Work')
      );

    if (shouldArchive) {
      file.moveTo(archiveFolder);
    }
  }
}

function getMostRecentTimecard(employeeFolder) {
  const files = employeeFolder.getFiles();

  let newestFile = null;
  let newestDate = null;

  while (files.hasNext()) {
    const file = files.next();

    if (file.getName().includes('Timecard')) {
      const updated = file.getLastUpdated();

      if (!newestDate || updated > newestDate) {
        newestDate = updated;
        newestFile = file;
      }
    }
  }

  return newestFile;
}

/************ DATE LOGIC ************/

function getCurrentRunInfo(referenceDate) {
  const anchorRunDate = new Date(2026, 4, 20);       // May 20, 2026 (Wed)
  const anchorPayDate = new Date(2026, 5, 5);        // June 5, 2026 (Fri)
  const anchorPeriodStart = new Date(2026, 4, 18);   // May 18, 2026 (Mon)
  const anchorPeriodEnd = new Date(2026, 4, 31);     // May 31, 2026 (Sun)

  const today = new Date(referenceDate || new Date());
  today.setHours(0, 0, 0, 0);

  const daysSinceAnchor = Math.floor(
    (today - anchorRunDate) / (1000 * 60 * 60 * 24)
  );

  const cycleOffset = Math.max(0, Math.floor(daysSinceAnchor / 14));

  const currentPeriodStart = addDays(anchorPeriodStart, cycleOffset * 14);
  const currentPeriodEnd = addDays(anchorPeriodEnd, cycleOffset * 14);
  const payDate = addDays(anchorPayDate, cycleOffset * 14);

  return {
    payDate: payDate,
    currentPeriodStart: currentPeriodStart,
    currentPeriodEnd: currentPeriodEnd,
    previousPeriodStart: addDays(currentPeriodStart, -14),
    previousPeriodEnd: addDays(currentPeriodEnd, -14)
  };
}

function isScheduledRunDay() {
  const anchorRunDate = new Date(2026, 4, 20); // Wednesday May 20, 2026

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const daysSinceAnchor = Math.floor(
    (today - anchorRunDate) / (1000 * 60 * 60 * 24)
  );

  return daysSinceAnchor >= 0 && daysSinceAnchor % 14 === 0;
}

/************ HELPERS ************/

function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  return parentFolder.createFolder(folderName);
}

function addDays(date, days) {
  const newDate = new Date(date);
  newDate.setDate(newDate.getDate() + days);
  return newDate;
}

function formatDateForFile(date) {
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'MM-dd-yyyy'
  );
}

function formatPeriodText(date) {
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'M/d'
  );
}

function formatPeriodFolderName(startDate, endDate, suffix) {
  return (
    formatPeriodText(startDate) +
    '-' +
    formatPeriodText(endDate) +
    suffix
  );
}

function fixTimecardProtections(spreadsheet) {
  const timecardSheet = spreadsheet.getSheetByName('TimeCardNotes');
  const inventorySheet = spreadsheet.getSheetByName('Inventory');

  const file = DriveApp.getFileById(spreadsheet.getId());
const editors = file.getEditors()
  .map(function(user) {
    return user.getEmail();
  })
  .filter(function(email) {
    return email && email.trim() !== '';
  });

  if (timecardSheet) {
    removeProtectionsFromSheet_(timecardSheet);
  }

  if (inventorySheet) {
    removeProtectionsFromSheet_(inventorySheet);

    protectInventoryRange_(inventorySheet, 'H:I', editors);
    protectInventoryRange_(inventorySheet, 'C:C', editors);
    protectInventoryRange_(inventorySheet, 'A:B', editors);
    protectInventoryRange_(inventorySheet, 'F:I', editors);
  }
}

function removeProtectionsFromSheet_(sheet) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p) {
    p.remove();
  });

  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) {
    p.remove();
  });
}

function protectInventoryRange_(sheet, rangeA1, editorEmails) {
  const protection = sheet.getRange(rangeA1).protect();

  protection.setDescription('Inventory!' + rangeA1);
  protection.removeEditors(protection.getEditors());

  if (editorEmails.length > 0) {
    protection.addEditors(editorEmails);
  }

  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}

function fixExistingTimecardSharingSettings() {
  const parentFolderIds = [
    '1abALtYC_Bcdnl-J06wtOxyJQI6XFdpQs',
    '1qF4Nv_MLLOEbd4i8nb6PF_2TsnryH0o2',
    '1B4ZbEOPkQ8GW-RETHpWLcQDob137SRhG',
    '1Y3kbrNn_v2E8oyJQYCiH0Swj8m2bmwna',
    '185swROseZQ4U1nRhHtD-pFNIRx0DIz19'
  ];

  parentFolderIds.forEach(function(parentFolderId) {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const employeeFolders = parentFolder.getFolders();

    while (employeeFolders.hasNext()) {
      const employeeFolder = employeeFolders.next();
      const files = employeeFolder.getFiles();

      while (files.hasNext()) {
        const file = files.next();

        if (!file.getName().includes('Timecard')) continue;

        try {
          const spreadsheet = SpreadsheetApp.openById(file.getId());
          fixTimecardProtections(spreadsheet);
          Logger.log('Updated protections: ' + file.getName());
        } catch (err) {
          Logger.log('ERROR updating ' + file.getName() + ': ' + err);
        }
      }
    }
  });
}
function convertExistingTimecardTimesFromUtcToCentral() {
  const parentFolderIds = [
    '1abALtYC_Bcdnl-J06wtOxyJQI6XFdpQs',
    '1qF4Nv_MLLOEbd4i8nb6PF_2TsnryH0o2',
    '1B4ZbEOPkQ8GW-RETHpWLcQDob137SRhG',
    '1Y3kbrNn_v2E8oyJQYCiH0Swj8m2bmwna',
    '185swROseZQ4U1nRhHtD-pFNIRx0DIz19'
  ];

  parentFolderIds.forEach(function(parentFolderId) {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const employeeFolders = parentFolder.getFolders();

    while (employeeFolders.hasNext()) {
      const employeeFolder = employeeFolders.next();
      const files = employeeFolder.getFiles();

      while (files.hasNext()) {
        const file = files.next();
        if (!file.getName().includes('Timecard')) continue;

        try {
          const ss = SpreadsheetApp.openById(file.getId());
          convertTimeCardNotesUtcToCentral_(ss);
          Logger.log('Converted UTC to Central: ' + file.getName());
        } catch (err) {
          Logger.log('ERROR converting ' + file.getName() + ': ' + err);
        }
      }
    }
  });
}

function convertTimeCardNotesUtcToCentral_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName('TimeCardNotes');
  if (!sheet) return;

  const firstDataRow = 2;
  const lastRow = sheet.getLastRow();
  if (lastRow < firstDataRow) return;

  const numRows = lastRow - firstDataRow + 1;

  const completedAtCol = 6;   // F
  const dateOnlyCol = 7;      // G
  const canceledAtCol = 8;    // H
  const weekendCol = 17;      // Q
  const afterHoursCol = 18;   // R

  const completedValues = sheet.getRange(firstDataRow, completedAtCol, numRows, 1).getValues();
  const canceledValues = sheet.getRange(firstDataRow, canceledAtCol, numRows, 1).getValues();
  const dateOnlyValues = sheet.getRange(firstDataRow, dateOnlyCol, numRows, 1).getValues();
const existingWeekendValues = sheet.getRange(firstDataRow, weekendCol, numRows, 1).getValues();
const existingAfterHoursValues = sheet.getRange(firstDataRow, afterHoursCol, numRows, 1).getValues();

  const newCompletedValues = [];
  const newDateOnlyValues = [];
  const newCanceledValues = [];
  const weekendValues = [];
  const afterHoursValues = [];

  for (let i = 0; i < numRows; i++) {
    const completedCentral = convertUtcDateToCentralDate_(completedValues[i][0]);
    const canceledCentral = convertUtcDateToCentralDate_(canceledValues[i][0]);

    newCompletedValues.push([completedCentral || completedValues[i][0]]);
    newCanceledValues.push([canceledCentral || canceledValues[i][0]]);

    if (completedCentral) {
      newDateOnlyValues.push([
        new Date(
          completedCentral.getFullYear(),
          completedCentral.getMonth(),
          completedCentral.getDate()
        )
      ]);

      const day = completedCentral.getDay();
      const hour = completedCentral.getHours();

      weekendValues.push([day === 0 || day === 6 ? 'Yes' : 'No']);
      afterHoursValues.push([hour >= 19 ? 'Yes' : 'No']);
} else {
  newDateOnlyValues.push([dateOnlyValues[i][0]]);
  weekendValues.push([existingWeekendValues[i][0]]);
  afterHoursValues.push([existingAfterHoursValues[i][0]]);
}
  }

  sheet.getRange(firstDataRow, completedAtCol, numRows, 1).setValues(newCompletedValues);
  sheet.getRange(firstDataRow, dateOnlyCol, numRows, 1).setValues(newDateOnlyValues);
  sheet.getRange(firstDataRow, canceledAtCol, numRows, 1).setValues(newCanceledValues);
  sheet.getRange(firstDataRow, weekendCol, numRows, 1).setValues(weekendValues);
  sheet.getRange(firstDataRow, afterHoursCol, numRows, 1).setValues(afterHoursValues);

  sheet.getRange(firstDataRow, completedAtCol, numRows, 1).setNumberFormat('M/d/yyyy h:mm AM/PM');
  sheet.getRange(firstDataRow, dateOnlyCol, numRows, 1).setNumberFormat('M/d/yyyy');
  sheet.getRange(firstDataRow, canceledAtCol, numRows, 1).setNumberFormat('M/d/yyyy h:mm AM/PM');
}

function convertUtcDateToCentralDate_(value) {
  if (!value) return null;

  let dateValue;

  if (value instanceof Date) {
    dateValue = new Date(value);
  } else {
    dateValue = new Date(value.toString());
  }

  if (isNaN(dateValue.getTime())) return null;

  dateValue.setHours(dateValue.getHours() - 5);

  return dateValue;
}
function forceSubtractFiveHoursOnOpenSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TimeCardNotes');

  const firstRow = 2;
  const lastRow = sheet.getLastRow();

  const fRange = sheet.getRange(firstRow, 6, lastRow - firstRow + 1, 1);
  const hRange = sheet.getRange(firstRow, 8, lastRow - firstRow + 1, 1);

  const fValues = fRange.getValues();
  const hValues = hRange.getValues();

  for (let i = 0; i < fValues.length; i++) {
    if (fValues[i][0] instanceof Date) {
      fValues[i][0] = new Date(fValues[i][0].getTime() - (5 * 60 * 60 * 1000));
    }

    if (hValues[i][0] instanceof Date) {
      hValues[i][0] = new Date(hValues[i][0].getTime() - (5 * 60 * 60 * 1000));
    }
  }

  fRange.setValues(fValues);
  hRange.setValues(hValues);
}
function forceSubtractFiveHoursOnAllTimecards() {
  const parentFolderIds = [
    '1abALtYC_Bcdnl-J06wtOxyJQI6XFdpQs',
    '1qF4Nv_MLLOEbd4i8nb6PF_2TsnryH0o2',
    '1B4ZbEOPkQ8GW-RETHpWLcQDob137SRhG',
    '1Y3kbrNn_v2E8oyJQYCiH0Swj8m2bmwna',
    '185swROseZQ4U1nRhHtD-pFNIRx0DIz19'
  ];

  parentFolderIds.forEach(function(parentFolderId) {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const employeeFolders = parentFolder.getFolders();

    while (employeeFolders.hasNext()) {
      const employeeFolder = employeeFolders.next();
      const files = employeeFolder.getFiles();

      while (files.hasNext()) {
        const file = files.next();
        if (!file.getName().includes('Timecard')) continue;

        try {
          const ss = SpreadsheetApp.openById(file.getId());
          forceSubtractFiveHoursFromSpreadsheet_(ss);
          Logger.log('Subtracted 5 hours: ' + file.getName());
        } catch (err) {
          Logger.log('ERROR subtracting 5 hours from ' + file.getName() + ': ' + err);
        }
      }
    }
  });
}

function forceSubtractFiveHoursFromSpreadsheet_(ss) {
  const sheet = ss.getSheetByName('TimeCardNotes');
  if (!sheet) return;

  const firstRow = 2;
  const lastRow = sheet.getLastRow();
  if (lastRow < firstRow) return;

  const fRange = sheet.getRange(firstRow, 6, lastRow - firstRow + 1, 1);
  const hRange = sheet.getRange(firstRow, 8, lastRow - firstRow + 1, 1);

  const fValues = fRange.getValues();
  const hValues = hRange.getValues();

  for (let i = 0; i < fValues.length; i++) {
    if (fValues[i][0] instanceof Date) {
      fValues[i][0] = new Date(fValues[i][0].getTime() - 5 * 60 * 60 * 1000);
    }

    if (hValues[i][0] instanceof Date) {
      hValues[i][0] = new Date(hValues[i][0].getTime() - 5 * 60 * 60 * 1000);
    }
  }

  fRange.setValues(fValues);
  hRange.setValues(hValues);
}
function forceSubtractFiveHoursOnProcessedPVFiles() {
  const parentFolderId = '185swROseZQ4U1nRhHtD-pFNIRx0DIz19';
  const parentFolder = DriveApp.getFolderById(parentFolderId);
  const employeeFolders = parentFolder.getFolders();

  while (employeeFolders.hasNext()) {
    const employeeFolder = employeeFolders.next();
    const files = employeeFolder.getFiles();

    while (files.hasNext()) {
      const file = files.next();

      if (!file.getName().includes('Processed')) continue;

      try {
        const ss = SpreadsheetApp.openById(file.getId());
        const sheet = ss.getSheetByName('Query result');
        if (!sheet) {
          Logger.log('No Query result tab: ' + file.getName());
          continue;
        }

        subtractFiveHoursFromQueryResult_(sheet);
        Logger.log('Updated Processed file: ' + file.getName());
      } catch (err) {
        Logger.log('ERROR updating ' + file.getName() + ': ' + err);
      }
    }
  }
}

function subtractFiveHoursFromQueryResult_(sheet) {
  const firstRow = 2;
  const lastRow = sheet.getLastRow();
  if (lastRow < firstRow) return;

  const completedAtCol = 6; // F
  const dateOnlyCol = 7;    // G
  const canceledAtCol = 8;  // H
  const weekendCol = 17;    // Q
  const afterHoursCol = 18; // R

  const numRows = lastRow - firstRow + 1;

  const fRange = sheet.getRange(firstRow, completedAtCol, numRows, 1);
  const hRange = sheet.getRange(firstRow, canceledAtCol, numRows, 1);

  const fValues = fRange.getValues();
  const hValues = hRange.getValues();

  const gValues = [];
  const qValues = [];
  const rValues = [];

  for (let i = 0; i < numRows; i++) {
    if (fValues[i][0] instanceof Date) {
      fValues[i][0] = new Date(fValues[i][0].getTime() - 5 * 60 * 60 * 1000);

      gValues.push([
        new Date(
          fValues[i][0].getFullYear(),
          fValues[i][0].getMonth(),
          fValues[i][0].getDate()
        )
      ]);

      const day = fValues[i][0].getDay();
      const hour = fValues[i][0].getHours();

      qValues.push([day === 0 || day === 6 ? 'Yes' : 'No']);
      rValues.push([hour >= 19 ? 'Yes' : 'No']);
    } else {
      gValues.push([sheet.getRange(firstRow + i, dateOnlyCol).getValue()]);
      qValues.push([sheet.getRange(firstRow + i, weekendCol).getValue()]);
      rValues.push([sheet.getRange(firstRow + i, afterHoursCol).getValue()]);
    }

    if (hValues[i][0] instanceof Date) {
      hValues[i][0] = new Date(hValues[i][0].getTime() - 5 * 60 * 60 * 1000);
    }
  }

  fRange.setValues(fValues);
  hRange.setValues(hValues);
  sheet.getRange(firstRow, dateOnlyCol, numRows, 1).setValues(gValues);
  sheet.getRange(firstRow, weekendCol, numRows, 1).setValues(qValues);
  sheet.getRange(firstRow, afterHoursCol, numRows, 1).setValues(rValues);

  fRange.setNumberFormat('M/d/yyyy h:mm AM/PM');
  hRange.setNumberFormat('M/d/yyyy h:mm AM/PM');
  sheet.getRange(firstRow, dateOnlyCol, numRows, 1).setNumberFormat('M/d/yyyy');
}
function convertProcessedXlsxFilesToGoogleSheets() {
  const parentFolderId = '185swROseZQ4U1nRhHtD-pFNIRx0DIz19';
  const parentFolder = DriveApp.getFolderById(parentFolderId);
  const employeeFolders = parentFolder.getFolders();

  while (employeeFolders.hasNext()) {
    const employeeFolder = employeeFolders.next();
    const files = employeeFolder.getFiles();

    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();

      if (!fileName.includes('Processed')) continue;
      if (!fileName.toLowerCase().endsWith('.xlsx')) continue;

      const newName = fileName.replace(/\.xlsx$/i, '');

      try {
        const convertedFile = Drive.Files.copy(
          {
            title: newName,
            mimeType: MimeType.GOOGLE_SHEETS
          },
          file.getId()
        );

        const newGoogleSheetFile = DriveApp.getFileById(convertedFile.id);
        employeeFolder.addFile(newGoogleSheetFile);

        // Optional: move original XLSX to trash after successful conversion
        // file.setTrashed(true);

        Logger.log('Converted to Google Sheet: ' + newName);
      } catch (err) {
        Logger.log('ERROR converting ' + fileName + ': ' + err);
      }
    }
  }
}
function trashProcessedXlsxFilesAfterConversion() {
  const parentFolderId = '185swROseZQ4U1nRhHtD-pFNIRx0DIz19';
  const parentFolder = DriveApp.getFolderById(parentFolderId);
  const employeeFolders = parentFolder.getFolders();

  while (employeeFolders.hasNext()) {
    const employeeFolder = employeeFolders.next();
    const files = employeeFolder.getFiles();

    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();

      if (!fileName.includes('Processed')) continue;
      if (!fileName.toLowerCase().endsWith('.xlsx')) continue;

      const googleSheetName = fileName.replace(/\.xlsx$/i, '');

      const matchingGoogleSheets = employeeFolder.getFilesByName(googleSheetName);

      if (matchingGoogleSheets.hasNext()) {
        file.setTrashed(true);
        Logger.log('Trashed XLSX duplicate: ' + fileName);
      }
    }
  }
}
function convertTopThreeProcessedXlsxFilesToGoogleSheets() {
  const parentFolderId = '1miO62yCilTD7CEiX1Y42beVEPXEdaVtM';
  const allowedFolderNames = [
    'DFW | Central TX | Steve L',
    'East TX | Steve K.',
    'Houston | SA'
  ];

  const parentFolder = DriveApp.getFolderById(parentFolderId);
  const folders = parentFolder.getFolders();

  while (folders.hasNext()) {
    const folder = folders.next();
    const folderName = folder.getName();

    if (!allowedFolderNames.includes(folderName)) continue;

    const files = folder.getFiles();

    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();

      if (!fileName.includes('Processed')) continue;
      if (!fileName.toLowerCase().endsWith('.xlsx')) continue;

      const newName = fileName.replace(/\.xlsx$/i, '');

      try {
        const convertedFile = Drive.Files.copy(
          {
            title: newName,
            mimeType: MimeType.GOOGLE_SHEETS
          },
          file.getId()
        );

        const newGoogleSheetFile = DriveApp.getFileById(convertedFile.id);
        folder.addFile(newGoogleSheetFile);

        Logger.log('Converted to Google Sheet: ' + folderName + ' / ' + newName);
      } catch (err) {
        Logger.log('ERROR converting ' + folderName + ' / ' + fileName + ': ' + err);
      }
    }
  }
}

function convertProcessedXlsxFilesInSpecificFolders() {
  const folderIds = [
    '1abALtYC_Bcdnl-J06wtOxyJQI6XFdpQs',
    '1qF4Nv_MLLOEbd4i8nb6PF_2TsnryH0o2',
    '1B4ZbEOPkQ8GW-RETHpWLcQDob137SRhG',
    '1Y3kbrNn_v2E8oyJQYCiH0Swj8m2bmwna'
  ];

  folderIds.forEach(function(folderId) {

    const regionFolder = DriveApp.getFolderById(folderId);
    const employeeFolders = regionFolder.getFolders();

    Logger.log('Checking region folder: ' + regionFolder.getName());

    while (employeeFolders.hasNext()) {

      const employeeFolder = employeeFolders.next();
      Logger.log('Employee folder: ' + employeeFolder.getName());

      const files = employeeFolder.getFiles();

      while (files.hasNext()) {

        const file = files.next();
        const fileName = file.getName();

        if (!fileName.includes('Processed')) continue;
        if (!fileName.toLowerCase().endsWith('.xlsx')) continue;

        const newName = fileName.replace(/\.xlsx$/i, '');

        try {

          const convertedFile = Drive.Files.copy(
            {
              title: newName,
              mimeType: MimeType.GOOGLE_SHEETS
            },
            file.getId()
          );

          const newGoogleSheet = DriveApp.getFileById(convertedFile.id);

          employeeFolder.addFile(newGoogleSheet);

          Logger.log('Converted: ' + newName);

        } catch (err) {

          Logger.log('ERROR converting ' + fileName + ': ' + err);

        }
      }
    }
  });
}
function trashDuplicateProcessedXlsxFilesInSpecificFolders() {
  const folderIds = [
    '1abALtYC_Bcdnl-J06wtOxyJQI6XFdpQs',
    '1qF4Nv_MLLOEbd4i8nb6PF_2TsnryH0o2',
    '1B4ZbEOPkQ8GW-RETHpWLcQDob137SRhG',
    '1Y3kbrNn_v2E8oyJQYCiH0Swj8m2bmwna'
  ];

  folderIds.forEach(function(folderId) {

    const regionFolder = DriveApp.getFolderById(folderId);
    const employeeFolders = regionFolder.getFolders();

    while (employeeFolders.hasNext()) {

      const employeeFolder = employeeFolders.next();
      const files = employeeFolder.getFiles();

      while (files.hasNext()) {

        const file = files.next();
        const fileName = file.getName();

        if (!fileName.includes('Processed')) continue;
        if (!fileName.toLowerCase().endsWith('.xlsx')) continue;

        const googleSheetName = fileName.replace(/\.xlsx$/i, '');

        const matchingGoogleSheets =
          employeeFolder.getFilesByName(googleSheetName);

        if (matchingGoogleSheets.hasNext()) {

          file.setTrashed(true);

          Logger.log('Trashed XLSX duplicate: ' + fileName);

        }
      }
    }
  });
}
function recalculateWeekendAndAfterHoursForPVProcessedFiles() {
  const parentFolderId = '185swROseZQ4U1nRhHtD-pFNIRx0DIz19';
  const parentFolder = DriveApp.getFolderById(parentFolderId);
  const employeeFolders = parentFolder.getFolders();

  while (employeeFolders.hasNext()) {
    const employeeFolder = employeeFolders.next();
    const files = employeeFolder.getFiles();

    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();

      if (!fileName.includes('Processed')) continue;
      if (fileName.toLowerCase().endsWith('.xlsx')) continue;

      try {
        const ss = SpreadsheetApp.openById(file.getId());
        const sheet = ss.getSheetByName('Query result');
        if (!sheet) continue;

        recalculateWeekendAndAfterHours_(sheet);
        Logger.log('Recalculated Q/R: ' + fileName);
      } catch (err) {
        Logger.log('ERROR recalculating ' + fileName + ': ' + err);
      }
    }
  }
}

function recalculateWeekendAndAfterHours_(sheet) {
  const firstRow = 2;
  const lastRow = sheet.getLastRow();
  if (lastRow < firstRow) return;

  const numRows = lastRow - firstRow + 1;

  const completedAtCol = 6;   // F
  const weekendCol = 17;      // Q
  const afterHoursCol = 18;   // R

  const displayValues = sheet.getRange(firstRow, completedAtCol, numRows, 1).getDisplayValues();

  const weekendValues = [];
  const afterHoursValues = [];

  for (let i = 0; i < numRows; i++) {
    const text = displayValues[i][0];

    if (!text) {
      weekendValues.push(['']);
      afterHoursValues.push(['']);
      continue;
    }

    const parsed = new Date(text);

    if (isNaN(parsed.getTime())) {
      weekendValues.push(['']);
      afterHoursValues.push(['']);
      continue;
    }

    const day = parsed.getDay();    // Sunday = 0, Saturday = 6
    const hour = parsed.getHours(); // 7 PM = 19

    weekendValues.push([day === 0 || day === 6 ? 'Yes' : 'No']);
    afterHoursValues.push([hour >= 19 ? 'Yes' : 'No']);
  }

  sheet.getRange(firstRow, weekendCol, numRows, 1).setValues(weekendValues);
  sheet.getRange(firstRow, afterHoursCol, numRows, 1).setValues(afterHoursValues);
}
function applyFolderEditorSharingToFile_(file, folder) {
  const editorEmails = folder.getEditors()
    .map(function(user) {
      return user.getEmail();
    })
    .filter(function(email) {
      return email && email.trim() !== '';
    });

  if (editorEmails.length > 0) {
    file.addEditors(editorEmails);
  }

  Logger.log('Added folder editors to file: ' + file.getName());
}
function lockCurrentPaydateTimecardsToViewOnly() {
  const parentFolderIds = [
    '1abALtYC_Bcdnl-J06wtOxyJQI6XFdpQs',
    '1qF4Nv_MLLOEbd4i8nb6PF_2TsnryH0o2',
    '1B4ZbEOPkQ8GW-RETHpWLcQDob137SRhG',
    '1Y3kbrNn_v2E8oyJQYCiH0Swj8m2bmwna',
    '185swROseZQ4U1nRhHtD-pFNIRx0DIz19'
  ];

  const runInfo = getCurrentRunInfo();

  const currentFileNamePart =
    'Timecard - Pay Date ' + formatDateForFile(runInfo.payDate);

  parentFolderIds.forEach(function(parentFolderId) {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const employeeFolders = parentFolder.getFolders();

    while (employeeFolders.hasNext()) {
      const employeeFolder = employeeFolders.next();
      const files = employeeFolder.getFiles();

      while (files.hasNext()) {
        const file = files.next();
        const fileName = file.getName();

        if (!fileName.includes(currentFileNamePart)) continue;

        try {
          changeFileEditorsToViewers_(file);
          Logger.log('Locked to view only: ' + fileName);
        } catch (err) {
          Logger.log('ERROR locking ' + fileName + ': ' + err);
        }
      }
    }
  });
}

function changeFileEditorsToViewers_(file) {
  const editors = file.getEditors();

  editors.forEach(function(user) {
    const email = user.getEmail();

    if (!email || email.trim() === '') return;

    file.removeEditor(email);
    file.addViewer(email);
  });
}
function calculateUniqueMileageByFacilityAndDate() {
  const ss = SpreadsheetApp.openById('PASTE_SPREADSHEET_ID_HERE');
  const sheet = ss.getSheetByName('Query result');

  const firstRow = 2;
  const lastRow = sheet.getLastRow();

  if (lastRow < firstRow) return;

  const numRows = lastRow - firstRow + 1;

  const facilityValues = sheet.getRange(firstRow, 3, numRows, 1).getDisplayValues(); // C
  const dateValues = sheet.getRange(firstRow, 7, numRows, 1).getDisplayValues();     // G
  const mileageValues = sheet.getRange(firstRow, 21, numRows, 1).getValues();        // U

  const output = [];
  const seen = {};

  for (let i = 0; i < numRows; i++) {

    const facility = facilityValues[i][0].trim();
    const date = dateValues[i][0].trim();
    const miles = mileageValues[i][0];

    const key = facility + '|' + date;

    if (!seen[key]) {
      seen[key] = true;
      output.push([miles]);
    } else {
      output.push([0]);
    }
  }

  sheet.getRange(firstRow, 23, numRows, 1).setValues(output); // W

  Logger.log('Updated unique mileage values in column W');
}
function calculateUniqueMileageForAllProcessedTimecards() {
  const folderIds = [
    '1abALtYC_Bcdnl-J06wtOxyJQI6XFdpQs',
    '1qF4Nv_MLLOEbd4i8nb6PF_2TsnryH0o2',
    '1B4ZbEOPkQ8GW-RETHpWLcQDob137SRhG',
    '1Y3kbrNn_v2E8oyJQYCiH0Swj8m2bmwna',
    '185swROseZQ4U1nRhHtD-pFNIRx0DIz19'
  ];

  folderIds.forEach(function(folderId) {
    const regionFolder = DriveApp.getFolderById(folderId);
    const employeeFolders = regionFolder.getFolders();

    while (employeeFolders.hasNext()) {
      const employeeFolder = employeeFolders.next();
      const files = employeeFolder.getFiles();

      while (files.hasNext()) {
        const file = files.next();
        const fileName = file.getName();

        if (!fileName.includes('Processed')) continue;
        if (fileName.toLowerCase().endsWith('.xlsx')) continue;

        try {
          const ss = SpreadsheetApp.openById(file.getId());
          const sheet = ss.getSheetByName('Query result');

          if (!sheet) {
            Logger.log('No Query result tab: ' + fileName);
            continue;
          }

          calculateUniqueMileageOnSheet_(sheet);
          Logger.log('Updated mileage: ' + fileName);

        } catch (err) {
          Logger.log('ERROR updating mileage for ' + fileName + ': ' + err);
        }
      }
    }
  });
}

function calculateUniqueMileageOnSheet_(sheet) {
  const firstRow = 2;
  const lastRow = sheet.getLastRow();

  if (lastRow < firstRow) return;

  const numRows = lastRow - firstRow + 1;

  const facilityValues = sheet.getRange(firstRow, 3, numRows, 1).getDisplayValues(); // C
  const dateValues = sheet.getRange(firstRow, 7, numRows, 1).getDisplayValues();     // G
  const mileageValues = sheet.getRange(firstRow, 21, numRows, 1).getValues();        // U

  const output = [];
  const seen = {};

  for (let i = 0; i < numRows; i++) {
    const facility = facilityValues[i][0].trim();
    const date = dateValues[i][0].trim();
    const miles = mileageValues[i][0];

    if (!facility || !date) {
      output.push(['']);
      continue;
    }

    const key = facility + '|' + date;

    if (!seen[key]) {
      seen[key] = true;
      output.push([miles]);
    } else {
      output.push([0]);
    }
  }

  sheet.getRange(firstRow, 23, numRows, 1).setValues(output); // W
}
function cleanupInventoryFormatting() {

  const folderIds = [
    '1abALtYC_Bcdnl-J06wtOxyJQI6XFdpQs',
    '1qF4Nv_MLLOEbd4i8nb6PF_2TsnryH0o2',
    '1B4ZbEOPkQ8GW-RETHpWLcQDob137SRhG',
    '1Y3kbrNn_v2E8oyJQYCiH0Swj8m2bmwna',
    '185swROseZQ4U1nRhHtD-pFNIRx0DIz19'
  ];

  folderIds.forEach(function(folderId) {

    const parentFolder = DriveApp.getFolderById(folderId);
    const employeeFolders = parentFolder.getFolders();

    while (employeeFolders.hasNext()) {

      const employeeFolder = employeeFolders.next();
      const files = employeeFolder.getFiles();

      while (files.hasNext()) {

        const file = files.next();

        if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

        try {

          const ss = SpreadsheetApp.openById(file.getId());
          const sheet = ss.getSheetByName('Inventory');

          if (!sheet) continue;

          // Fix merged header area
          sheet.getRange('C1:F1').breakApart();
          sheet.getRange('C1:F1').merge();

          // Clean formatting
          sheet.getRange('1:2')
            .setWrap(false)
            .setVerticalAlignment('middle')
            .setHorizontalAlignment('center');

          // Fix row heights
          sheet.setRowHeight(1, 30);
          sheet.setRowHeight(2, 28);

          // Optional font cleanup
          sheet.getRange('1:2')
            .setFontSize(10)
            .setFontWeight('bold');

          Logger.log('Cleaned Inventory formatting: ' + file.getName());

        } catch (err) {

          Logger.log('ERROR cleaning ' + file.getName() + ': ' + err);

        }
      }
    }
  });
}
 function fixInfusionInventoryBanner_(spreadsheet, bannerDateText) {
  const sheet = spreadsheet.getSheetByName('Inventory');
  if (!sheet) return;

  sheet.getRange('A1:I1').breakApart();
  sheet.getRange('A1:I1').clearContent();

  const banner = sheet.getRange('B1:H1');
  banner.merge();

  banner.setValue('INVENTORY COUNTS MUST BE COMPLETED BY EOD ON ' + bannerDateText);

  banner
    .setFontWeight('bold')
    .setFontSize(10)
    .setWrap(false)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#FFF2CC');

  sheet.setRowHeight(1, 28);
}
function fixVascularInventoryBanner_(spreadsheet, bannerDateText) {
  const sheet = spreadsheet.getSheetByName('Inventory');
  if (!sheet) return;

  sheet.getRange('A1:K1').breakApart();
  sheet.getRange('A1:K1').clearContent();

  const banner = sheet.getRange('B1:H1');
  banner.merge();

  banner.setValue('INVENTORY COUNTS MUST BE COMPLETED BY EOD ON ' + bannerDateText);

  banner
    .setFontWeight('bold')
    .setFontSize(10)
    .setWrap(false)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#FFF2CC');

  sheet.setRowHeight(1, 28);
}

function cleanupVascularInventoryFormatting() {
  const bannerDateText = '5/17';

  const parentFolderId = '185swROseZQ4U1nRhHtD-pFNIRx0DIz19';
  const parentFolder = DriveApp.getFolderById(parentFolderId);
  const employeeFolders = parentFolder.getFolders();

  while (employeeFolders.hasNext()) {
    const employeeFolder = employeeFolders.next();
    const files = employeeFolder.getFiles();

    while (files.hasNext()) {
      const file = files.next();

      if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;
      if (!file.getName().includes('Timecard')) continue;

      const ss = SpreadsheetApp.openById(file.getId());
      fixVascularInventoryBanner_(ss, bannerDateText);

      Logger.log('Fixed vascular: ' + file.getName());
    }
  }
}

function cleanupInfusionInventoryFormatting() {
  const bannerDateText = '5/17';

  const parentFolderIds = [
    '1abALtYC_Bcdnl-J06wtOxyJQI6XFdpQs',
    '1qF4Nv_MLLOEbd4i8nb6PF_2TsnryH0o2',
    '1B4ZbEOPkQ8GW-RETHpWLcQDob137SRhG',
    '1Y3kbrNn_v2E8oyJQYCiH0Swj8m2bmwna'
  ];

  parentFolderIds.forEach(function(parentFolderId) {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const employeeFolders = parentFolder.getFolders();

    while (employeeFolders.hasNext()) {
      const employeeFolder = employeeFolders.next();
      const files = employeeFolder.getFiles();

      while (files.hasNext()) {
        const file = files.next();

        if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;
        if (!file.getName().includes('Timecard')) continue;

        const ss = SpreadsheetApp.openById(file.getId());
        fixInfusionInventoryBanner_(ss, bannerDateText);

        Logger.log('Fixed infusion: ' + file.getName());
      }
    }
  });
}
function forceFixInventoryBannersTo517() {
  forceFixInfusionInventoryBannersTo517_();
  forceFixVascularInventoryBannersTo517_();
}

function forceFixInfusionInventoryBannersTo517_() {
  const parentFolderIds = [
    '1abALtYC_Bcdnl-J06wtOxyJQI6XFdpQs',
    '1qF4Nv_MLLOEbd4i8nb6PF_2TsnryH0o2',
    '1B4ZbEOPkQ8GW-RETHpWLcQDob137SRhG',
    '1Y3kbrNn_v2E8oyJQYCiH0Swj8m2bmwna'
  ];

  parentFolderIds.forEach(function(parentFolderId) {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const employeeFolders = parentFolder.getFolders();

    while (employeeFolders.hasNext()) {
      const employeeFolder = employeeFolders.next();
      const files = employeeFolder.getFiles();

      while (files.hasNext()) {
        const file = files.next();
        if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;
        if (!file.getName().includes('Timecard')) continue;

        const ss = SpreadsheetApp.openById(file.getId());
        const sheet = ss.getSheetByName('Inventory');
        if (!sheet) continue;

        sheet.getRange('A1:I1').breakApart();
        sheet.getRange('A1:I1').clearContent();

        const banner = sheet.getRange('B1:H1');
        banner.merge();
        banner.setValue('INVENTORY COUNTS MUST BE COMPLETED BY EOD ON 5/17');
        banner.setFontWeight('bold').setFontSize(10).setWrap(false)
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle')
          .setBackground('#FFF2CC');

        sheet.setRowHeight(1, 28);
        Logger.log('Fixed infusion banner: ' + file.getName());
      }
    }
  });
}

function forceFixVascularInventoryBannersTo517_() {
  const parentFolderId = '185swROseZQ4U1nRhHtD-pFNIRx0DIz19';
  const parentFolder = DriveApp.getFolderById(parentFolderId);
  const employeeFolders = parentFolder.getFolders();

  while (employeeFolders.hasNext()) {
    const employeeFolder = employeeFolders.next();
    const files = employeeFolder.getFiles();

    while (files.hasNext()) {
      const file = files.next();
      if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;
      if (!file.getName().includes('Timecard')) continue;

      const ss = SpreadsheetApp.openById(file.getId());
      const sheet = ss.getSheetByName('Inventory');
      if (!sheet) continue;

      sheet.getRange('A1:K1').breakApart();
      sheet.getRange('A1:K1').clearContent();

      const banner = sheet.getRange('B1:H1');
      banner.merge();
      banner.setValue('INVENTORY COUNTS MUST BE COMPLETED BY EOD ON 5/17');
      banner.setFontWeight('bold').setFontSize(10).setWrap(false)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setBackground('#FFF2CC');

      sheet.setRowHeight(1, 28);
      Logger.log('Fixed vascular banner: ' + file.getName());
    }
  }
}
function forceFixInventoryBannersTo517_SAFE() {
  fixInventoryBannersInFolders_(
    [
      '1abALtYC_Bcdnl-J06wtOxyJQI6XFdpQs',
      '1qF4Nv_MLLOEbd4i8nb6PF_2TsnryH0o2',
      '1B4ZbEOPkQ8GW-RETHpWLcQDob137SRhG',
      '1Y3kbrNn_v2E8oyJQYCiH0Swj8m2bmwna'
    ],
    'B1:H1',
    'Fixed infusion banner: '
  );

  fixInventoryBannersInFolders_(
    ['185swROseZQ4U1nRhHtD-pFNIRx0DIz19'],
    'B1:H1',
    'Fixed vascular banner: '
  );
}

function fixInventoryBannersInFolders_(parentFolderIds, bannerRangeA1, logPrefix) {
  parentFolderIds.forEach(function(parentFolderId) {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const employeeFolders = parentFolder.getFolders();

    while (employeeFolders.hasNext()) {
      const employeeFolder = employeeFolders.next();
      const files = employeeFolder.getFiles();

      while (files.hasNext()) {
        const file = files.next();

        if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;
        if (!file.getName().includes('Timecard')) continue;
        if (!file.getName().includes('05-22-2026')) continue;

        try {
          const ss = SpreadsheetApp.openById(file.getId());
          const sheet = ss.getSheetByName('Inventory');
          if (!sheet) continue;

          fixOneInventoryBanner517_(sheet, bannerRangeA1);

          Logger.log(logPrefix + file.getName());
        } catch (err) {
          Logger.log('ERROR fixing banner for ' + file.getName() + ': ' + err);
        }
      }
    }
  });
}

function fixOneInventoryBanner517_(sheet, bannerRangeA1) {
  const fullTopRow = sheet.getRange('A1:K1');

  fullTopRow.getMergedRanges().forEach(function(range) {
    range.breakApart();
  });

  fullTopRow.clearContent();

  const banner = sheet.getRange(bannerRangeA1);

  banner.getMergedRanges().forEach(function(range) {
    range.breakApart();
  });

  banner.merge();
  banner.setValue('INVENTORY COUNTS MUST BE COMPLETED BY EOD ON 5/17');

  banner
    .setFontWeight('bold')
    .setFontSize(10)
    .setWrap(false)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#FFF2CC');

  sheet.setRowHeight(1, 28);
}
function forceWriteInventoryBannerTextOnly() {
  const bannerText = 'INVENTORY COUNTS MUST BE COMPLETED BY EOD ON 5/17';

  const parentFolderIds = [
    '1abALtYC_Bcdnl-J06wtOxyJQI6XFdpQs',
    '1qF4Nv_MLLOEbd4i8nb6PF_2TsnryH0o2',
    '1B4ZbEOPkQ8GW-RETHpWLcQDob137SRhG',
    '1Y3kbrNn_v2E8oyJQYCiH0Swj8m2bmwna',
    '185swROseZQ4U1nRhHtD-pFNIRx0DIz19'
  ];

  parentFolderIds.forEach(function(parentFolderId) {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const employeeFolders = parentFolder.getFolders();

    while (employeeFolders.hasNext()) {
      const employeeFolder = employeeFolders.next();
      const files = employeeFolder.getFiles();

      while (files.hasNext()) {
        const file = files.next();

        if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;
        if (!file.getName().includes('Timecard')) continue;

        const ss = SpreadsheetApp.openById(file.getId());
        const sheet = ss.getSheetByName('Inventory');
        if (!sheet) continue;

        sheet.getRange('B1').setValue(bannerText);
        sheet.getRange('B1:H1')
          .setFontWeight('bold')
          .setFontSize(10)
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle')
          .setBackground('#FFF2CC');

        Logger.log('Wrote banner text: ' + file.getName());
      }
    }
  });
}
function updateCancelledOnSitePayForVascularProcessedSheets() {
  const folderId = '185swROseZQ4U1nRhHtD-pFNIRx0DIz19';
  const parentFolder = DriveApp.getFolderById(folderId);
  const employeeFolders = parentFolder.getFolders();

  while (employeeFolders.hasNext()) {
    const employeeFolder = employeeFolders.next();
    const files = employeeFolder.getFiles();

    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();

      if (!fileName.includes('Processed')) continue;
      if (fileName.toLowerCase().endsWith('.xlsx')) continue;
      if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

      try {
        const ss = SpreadsheetApp.openById(file.getId());
        updateCancelledOnSitePay_(ss);
        Logger.log('Updated cancelled on site: ' + fileName);
      } catch (err) {
        Logger.log('ERROR updating ' + fileName + ': ' + err);
      }
    }
  }
}

function updateCancelledOnSitePay_(ss) {
  const summary = ss.getSheetByName('Summary');
  const query = ss.getSheetByName('Query result');

  if (!summary || !query) return;

  const firstRow = 2;
  const lastRow = query.getLastRow();
  if (lastRow < firstRow) return;

  const numRows = lastRow - firstRow + 1;

  const cancelLocationValues = query.getRange(firstRow, 10, numRows, 1).getDisplayValues(); // J
  const procedureValues = query.getRange(firstRow, 13, numRows, 1).getDisplayValues(); // M

  const summaryProcedures = summary.getRange(2, 1, 20, 1).getDisplayValues(); // A
  const summaryPays = summary.getRange(2, 3, 20, 1).getDisplayValues(); // C

  const payLookup = {};

  for (let i = 0; i < summaryProcedures.length; i++) {
    const procedureName = summaryProcedures[i][0].toString().trim().toLowerCase();
    const normalPay = parseMoney_(summaryPays[i][0]);

    if (procedureName && normalPay > 0) {
      payLookup[procedureName] = normalPay;
    }
  }

  let cancelledOnSiteCount = 0;
  let cancelledOnSitePay = 0;

  for (let i = 0; i < numRows; i++) {
    const cancelLocation = cancelLocationValues[i][0].toString().toLowerCase();
    const procedure = procedureValues[i][0].toString().trim().toLowerCase();

    const isOnSite =
      cancelLocation.includes('on_site') ||
      cancelLocation.includes('on site') ||
      cancelLocation.includes('onsite');

    if (!isOnSite) continue;

    const normalPay = payLookup[procedure] || 0;

    cancelledOnSiteCount++;
    cancelledOnSitePay += normalPay * 0.5;
  }

  summary.getRange('A15').setValue('Cancelled on site');
  summary.getRange('B15').setValue(cancelledOnSiteCount);
  summary.getRange('C15').setValue(cancelledOnSitePay);
  summary.getRange('C15').setNumberFormat('$#,##0.00');
}

function parseMoney_(value) {
  if (!value) return 0;

  const cleaned = value
    .toString()
    .replace(/\$/g, '')
    .replace(/,/g, '')
    .trim();

  const number = Number(cleaned);

  return isNaN(number) ? 0 : number;
}
function fixVascularMileageSummaryAndRemoveQuerySum() {
  const folderId = '185swROseZQ4U1nRhHtD-pFNIRx0DIz19';
  const parentFolder = DriveApp.getFolderById(folderId);
  const employeeFolders = parentFolder.getFolders();

  while (employeeFolders.hasNext()) {
    const employeeFolder = employeeFolders.next();
    const files = employeeFolder.getFiles();

    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();

      if (!fileName.includes('Processed')) continue;
      if (fileName.toLowerCase().endsWith('.xlsx')) continue;
      if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

      try {
        const ss = SpreadsheetApp.openById(file.getId());
        const summary = ss.getSheetByName('Summary');
        const query = ss.getSheetByName('Query result');

        if (!summary || !query) continue;

        const lastRow = query.getLastRow();

        // Remove any SUM formula in column W on Query result
        const wRange = query.getRange(2, 23, Math.max(lastRow - 1, 1), 1);
        const formulas = wRange.getFormulas();

        for (let i = 0; i < formulas.length; i++) {
          if (formulas[i][0] && formulas[i][0].toUpperCase().includes('SUM')) {
            query.getRange(i + 2, 23).clearContent();
          }
        }

        // Set Summary B28 to equal sum of Query result column W values
        summary.getRange('B28').setFormula("=SUM('Query result'!W:W)");

        Logger.log('Fixed mileage summary: ' + fileName);
      } catch (err) {
        Logger.log('ERROR fixing mileage summary for ' + fileName + ': ' + err);
      }
    }
  }
}
function buildMasterPayrollSummary() {
  const masterSpreadsheetId = '1xnhAbP3TGVdZ62fvrmWO6rNlYY8gcMJQL6pyUh9DuzM';

  const infusionFolderIds = [
    '1abALtYC_Bcdnl-J06wtOxyJQI6XFdpQs',
    '1qF4Nv_MLLOEbd4i8nb6PF_2TsnryH0o2',
    '1B4ZbEOPkQ8GW-RETHpWLcQDob137SRhG',
    '1Y3kbrNn_v2E8oyJQYCiH0Swj8m2bmwna'
  ];

  const vascularFolderIds = [
    '185swROseZQ4U1nRhHtD-pFNIRx0DIz19'
  ];

  const master = SpreadsheetApp.openById(masterSpreadsheetId);

  const vascularTab = master.getSheetByName('Vascular');
  const infusionTab = master.getSheetByName('Infusion');

  if (!vascularTab || !infusionTab) {
    throw new Error('Master sheet must have tabs named Vascular and Infusion.');
  }

  clearMasterRows_(vascularTab, 2, 7);
  clearMasterRows_(infusionTab, 2, 4);

  const vascularRows = collectVascularSummaryRows_(vascularFolderIds);
  const infusionRows = collectInfusionSummaryRows_(infusionFolderIds);

  if (vascularRows.length > 0) {
    vascularTab.getRange(2, 1, vascularRows.length, 7).setValues(vascularRows);
  }

  if (infusionRows.length > 0) {
    infusionTab.getRange(2, 1, infusionRows.length, 4).setValues(infusionRows);
  }

  Logger.log('Master payroll summary rebuilt.');
}

function collectVascularSummaryRows_(folderIds) {
  const rows = [];

  folderIds.forEach(function(folderId) {
    const parentFolder = DriveApp.getFolderById(folderId);
    const employeeFolders = parentFolder.getFolders();

    while (employeeFolders.hasNext()) {
      const employeeFolder = employeeFolders.next();
      const files = employeeFolder.getFiles();

      while (files.hasNext()) {
        const file = files.next();
        const fileName = file.getName();

        if (!fileName.includes('Processed')) continue;
        if (fileName.toLowerCase().endsWith('.xlsx')) continue;
        if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

        try {
          const ss = SpreadsheetApp.openById(file.getId());
          const summary = ss.getSheetByName('Summary');
          if (!summary) continue;

          const employeeName = employeeFolder.getName();

          rows.push([
            employeeName,
            summary.getRange('B16').getValue(),
            summary.getRange('B18').getValue(),
            summary.getRange('B22').getValue(),
            summary.getRange('B26').getValue(),
            '',
            summary.getRange('B30').getValue()
          ]);

          Logger.log('Added vascular summary: ' + employeeName);
        } catch (err) {
          Logger.log('ERROR reading vascular file ' + fileName + ': ' + err);
        }
      }
    }
  });

  return rows;
}
function buildMasterPayrollSummary() {
  const masterSpreadsheetId = '1xnhAbP3TGVdZ62fvrmWO6rNlYY8gcMJQL6pyUh9DuzM';

  const infusionFolderIds = [
    '1abALtYC_Bcdnl-J06wtOxyJQI6XFdpQs',
    '1qF4Nv_MLLOEbd4i8nb6PF_2TsnryH0o2',
    '1B4ZbEOPkQ8GW-RETHpWLcQDob137SRhG',
    '1Y3kbrNn_v2E8oyJQYCiH0Swj8m2bmwna'
  ];

  const vascularFolderIds = [
    '185swROseZQ4U1nRhHtD-pFNIRx0DIz19'
  ];

  const master = SpreadsheetApp.openById(masterSpreadsheetId);

  const vascularTab = master.getSheetByName('Vascular');
  const infusionTab = master.getSheetByName('Infusion');

  vascularTab.getRange('A2:G').clearContent();
  infusionTab.getRange('A2:F').clearContent();

  const vascularRows = collectVascularSummaryRows_(vascularFolderIds);
  const infusionRows = collectInfusionSummaryRows_(infusionFolderIds);

  if (vascularRows.length > 0) {
    vascularTab.getRange(2, 1, vascularRows.length, 7).setValues(vascularRows);
  }

  if (infusionRows.length > 0) {
    infusionTab.getRange(2, 1, infusionRows.length, 6).setValues(infusionRows);
  }

  Logger.log('Master payroll summary rebuilt.');
}

function collectVascularSummaryRows_(folderIds) {
  const rows = [];

  folderIds.forEach(function(folderId) {
    const parentFolder = DriveApp.getFolderById(folderId);
    const employeeFolders = parentFolder.getFolders();

    while (employeeFolders.hasNext()) {
      const employeeFolder = employeeFolders.next();
      const employeeName = employeeFolder.getName();
      const files = employeeFolder.getFiles();

      while (files.hasNext()) {
        const file = files.next();
        const fileName = file.getName();

        if (!fileName.includes('Processed')) continue;
        if (fileName.toLowerCase().endsWith('.xlsx')) continue;
        if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

        const ss = SpreadsheetApp.openById(file.getId());
        const summary = ss.getSheetByName('Summary');
        if (!summary) continue;

        rows.push([
          employeeName,                    // A Employee Name
          summary.getRange('B16').getValue(), // B Line Total
          summary.getRange('B18').getValue(), // C Total Points
          summary.getRange('B22').getValue(), // D Weekend Pay
          summary.getRange('B26').getValue(), // E After hours pay
          '',                              // F Bonus
          summary.getRange('B30').getValue()  // G Mileage
        ]);
      }
    }
  });

  return rows;
}

function collectInfusionSummaryRows_(folderIds) {
  const rows = [];

  folderIds.forEach(function(folderId) {
    const parentFolder = DriveApp.getFolderById(folderId);
    const employeeFolders = parentFolder.getFolders();

    while (employeeFolders.hasNext()) {
      const employeeFolder = employeeFolders.next();
      const employeeName = employeeFolder.getName();
      const files = employeeFolder.getFiles();

      while (files.hasNext()) {
        const file = files.next();
        const fileName = file.getName();

        if (!fileName.includes('Processed')) continue;
        if (fileName.toLowerCase().endsWith('.xlsx')) continue;
        if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

        const ss = SpreadsheetApp.openById(file.getId());
        const summary = ss.getSheetByName('Summary');
        if (!summary) continue;

        rows.push([
          employeeName,                    // A Employee Name
          summary.getRange('B1').getValue(),  // B Total Infusions Completed
          summary.getRange('B5').getValue(),  // C Total Infusion Pay
          '',                              // D Total Transfusions
          '',                              // E Transfusion Pay
          summary.getRange('B14').getValue()  // F Mile Reimb
        ]);
      }
    }
  });

  return rows;
}
function archiveProcessedFilesOnly() {
  const archiveFolderNameInfusion = '4/20-5/3_Time_Cards';
  const archiveFolderNameVascular = '4/20-5/3_PV';

  const infusionAdminFolder = DriveApp.getFolderById('11e24KTEMKtPKEqDKY8gdwp5hsB6H1Wp1');
  const vascularAdminFolder = DriveApp.getFolderById('1kF1oZsaaXQGUNArRM58xp1Ac7tqLIDzG');

  const infusionArchiveFolder = getOrCreateFolderLocal_(infusionAdminFolder, archiveFolderNameInfusion);
  const vascularArchiveFolder = getOrCreateFolderLocal_(vascularAdminFolder, archiveFolderNameVascular);

  archiveProcessedInFolders_(
    [
      '1abALtYC_Bcdnl-J06wtOxyJQI6XFdpQs',
      '1qF4Nv_MLLOEbd4i8nb6PF_2TsnryH0o2',
      '1B4ZbEOPkQ8GW-RETHpWLcQDob137SRhG',
      '1Y3kbrNn_v2E8oyJQYCiH0Swj8m2bmwna'
    ],
    infusionArchiveFolder
  );

  archiveProcessedInFolders_(
    ['185swROseZQ4U1nRhHtD-pFNIRx0DIz19'],
    vascularArchiveFolder
  );
}

function getOrCreateFolderLocal_(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(folderName);
}

function archiveProcessedInFolders_(parentFolderIds, archiveFolder) {
  parentFolderIds.forEach(function(parentFolderId) {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const employeeFolders = parentFolder.getFolders();

    while (employeeFolders.hasNext()) {
      const employeeFolder = employeeFolders.next();
      const files = employeeFolder.getFiles();

      while (files.hasNext()) {
        const file = files.next();
        const fileName = file.getName();

        if (!fileName.includes('Processed')) continue;

        try {
          file.moveTo(archiveFolder);
          Logger.log('Archived Processed file: ' + fileName);
        } catch (err) {
          Logger.log('ERROR archiving ' + fileName + ': ' + err);
        }
      }
    }
  });
}
function testClickUpConnection() {
  const token = PropertiesService.getScriptProperties().getProperty('CLICKUP_API_TOKEN');

  if (!token) {
    Logger.log('ERROR: No ClickUp API token found. Set CLICKUP_API_TOKEN in Script Properties.');
    return;
  }

  const url = 'https://api.clickup.com/api/v2/user';

  const options = {
    method: 'get',
    headers: {
      Authorization: token
    },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    const body = response.getContentText();

    if (statusCode !== 200) {
      Logger.log('ERROR: ClickUp API returned status ' + statusCode + ': ' + body);
      return;
    }

    const user = JSON.parse(body).user;

    if (!user) {
      Logger.log('ERROR: Unexpected response from ClickUp: ' + body);
      return;
    }

    Logger.log('ClickUp connection successful. Username: ' + user.username + ', Email: ' + user.email);
  } catch (err) {
    Logger.log('ERROR: Failed to connect to ClickUp: ' + err);
  }
}
function syncPrimaryServiceLineFromDriveFolders() {
  const DRY_RUN = true;

  const CLICKUP_LIST_ID = '901711419133';
  const PRIMARY_SERVICE_LINE_FIELD_ID = 'c7918543-bdfd-4d16-8e6c-6c587b498bc0';

  const OPTION_IDS = {
    Infusion: 'fed98116-a3ca-44e0-870d-f53fdedbb98f',
    Vascular: '2eb1aecd-0cbf-4db6-9321-45823b151146',
    Corporate: '25298b53-d4ad-4d0c-afa4-616a63a1a9fa',
    'Needs Review': '4f8d8401-2a92-49bc-954e-a0432aeba457'
  };

  const FOLDER_GROUPS = {
    Infusion: [
      '1abALtYC_Bcdnl-J06wtOxyJQI6XFdpQs',
      '1qF4Nv_MLLOEbd4i8nb6PF_2TsnryH0o2',
      '1B4ZbEOPkQ8GW-RETHpWLcQDob137SRhG',
      '1Y3kbrNn_v2E8oyJQYCiH0Swj8m2bmwna'
    ],
    Vascular: ['185swROseZQ4U1nRhHtD-pFNIRx0DIz19'],
    Corporate: ['1UfIBI1wHcMbV_cz9KZSVfI8ud933gOBO']
  };

  const token = PropertiesService.getScriptProperties().getProperty('CLICKUP_API_TOKEN');

  if (!token) {
    Logger.log('ERROR: No ClickUp API token found. Set CLICKUP_API_TOKEN in Script Properties.');
    return;
  }

  function normalizeName(name) {
    let n = String(name).toLowerCase();

    // Strip a trailing employment-status suffix first (underscore- or space-separated)
    n = n.replace(/[_ ](ft|prn|pt|1099)$/i, '');

    // Strip a leading numeric prefix like "1. ", "2. ", "3. "
    n = n.replace(/^\d+\.\s*/, '');

    // Remove any text inside parentheses
    n = n.replace(/\([^)]*\)/g, ' ');

    // Treat commas, underscores, periods, and stray parentheses as whitespace
    n = n.replace(/[,._()]/g, ' ');

    // Collapse whitespace and trim
    n = n.replace(/\s+/g, ' ').trim();

    // Strip trailing FT/PRN/PT/1099 whole-word suffixes
    let prev;
    do {
      prev = n;
      n = n.replace(/\s+(ft|prn|pt|1099)$/i, '').trim();
    } while (n !== prev);

    // Sort words alphabetically so word order doesn't matter
    n = n.split(/\s+/).sort().join(' ');

    return n;
  }

  // Build lookup of normalized employee folder name -> service line
  const nameToServiceLine = {};

  Object.keys(FOLDER_GROUPS).forEach(function(serviceLine) {
    FOLDER_GROUPS[serviceLine].forEach(function(folderId) {
      const parentFolder = DriveApp.getFolderById(folderId);
      const employeeFolders = parentFolder.getFolders();

      while (employeeFolders.hasNext()) {
        const employeeFolder = employeeFolders.next();
        const key = normalizeName(employeeFolder.getName());

        if (!key || key.includes('archive')) {
          continue;
        }

        nameToServiceLine[key] = serviceLine;
      }
    });
  });

  // Fetch every task from the list, paging until a page returns no tasks
  const allTasks = [];
  let page = 0;

  while (true) {
    const listUrl =
      'https://api.clickup.com/api/v2/list/' + CLICKUP_LIST_ID +
      '/task?include_closed=true&subtasks=true&page=' + page;

    const response = UrlFetchApp.fetch(listUrl, {
      method: 'get',
      headers: { Authorization: token },
      muteHttpExceptions: true
    });

    const statusCode = response.getResponseCode();
    const body = response.getContentText();

    if (statusCode !== 200) {
      Logger.log('ERROR: Failed to fetch tasks (page ' + page + '), status ' + statusCode + ': ' + body);
      return;
    }

    const tasks = JSON.parse(body).tasks || [];

    if (tasks.length === 0) break;

    tasks.forEach(function(task) {
      allTasks.push(task);
    });

    page++;
  }

  Logger.log('Fetched ' + allTasks.length + ' tasks from ClickUp.');

  const counts = { Infusion: 0, Vascular: 0, Corporate: 0, 'Needs Review': 0 };

  allTasks.forEach(function(task) {
    if (task.parent) {
      Logger.log('SKIP subtask: ' + task.name);
      return;
    }

    const taskName = task.name;
    const key = normalizeName(taskName);
    const serviceLine = nameToServiceLine[key] || 'Needs Review';
    const optionId = OPTION_IDS[serviceLine];

    try {
      if (DRY_RUN) {
        counts[serviceLine]++;
        Logger.log('DRY RUN: would update "' + taskName + '" -> ' + serviceLine);
        return;
      }

      const updateUrl =
        'https://api.clickup.com/api/v2/task/' + task.id +
        '/field/' + PRIMARY_SERVICE_LINE_FIELD_ID;

      const updateResponse = UrlFetchApp.fetch(updateUrl, {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: token },
        payload: JSON.stringify({ value: optionId }),
        muteHttpExceptions: true
      });

      const code = updateResponse.getResponseCode();

      if (code === 200) {
        counts[serviceLine]++;
        Logger.log('Updated "' + taskName + '" -> ' + serviceLine);
      } else {
        Logger.log('ERROR updating "' + taskName + '" (status ' + code + '): ' + updateResponse.getContentText());
      }
    } catch (err) {
      Logger.log('ERROR updating "' + taskName + '": ' + err);
    }
  });

  Logger.log(
    'Done. Infusion: ' + counts.Infusion +
    ', Vascular: ' + counts.Vascular +
    ', Corporate: ' + counts.Corporate +
    ', Needs Review: ' + counts['Needs Review']
  );
}
function clearPrimaryServiceLineFromSubtasks() {
  const DRY_RUN = true;

  const CLICKUP_LIST_ID = '901711419133';
  const PRIMARY_SERVICE_LINE_FIELD_ID = 'c7918543-bdfd-4d16-8e6c-6c587b498bc0';

  const token = PropertiesService.getScriptProperties().getProperty('CLICKUP_API_TOKEN');

  if (!token) {
    Logger.log('ERROR: No ClickUp API token found. Set CLICKUP_API_TOKEN in Script Properties.');
    return;
  }

  // Fetch every task from the list, paging until a page returns no tasks
  const allTasks = [];
  let page = 0;

  while (true) {
    const listUrl =
      'https://api.clickup.com/api/v2/list/' + CLICKUP_LIST_ID +
      '/task?include_closed=true&subtasks=true&page=' + page;

    const response = UrlFetchApp.fetch(listUrl, {
      method: 'get',
      headers: { Authorization: token },
      muteHttpExceptions: true
    });

    const statusCode = response.getResponseCode();
    const body = response.getContentText();

    if (statusCode !== 200) {
      Logger.log('ERROR: Failed to fetch tasks (page ' + page + '), status ' + statusCode + ': ' + body);
      return;
    }

    const tasks = JSON.parse(body).tasks || [];

    if (tasks.length === 0) break;

    tasks.forEach(function(task) {
      allTasks.push(task);
    });

    page++;
  }

  Logger.log('Fetched ' + allTasks.length + ' tasks from ClickUp.');

  let clearedCount = 0;

  allTasks.forEach(function(task) {
    if (!task.parent) return;

    const taskName = task.name;

    try {
      if (DRY_RUN) {
        Logger.log('DRY RUN - would clear Primary Service Line on subtask "' + taskName + '" (' + task.id + ')');
        clearedCount++;
        return;
      }

      const clearUrl =
        'https://api.clickup.com/api/v2/task/' + task.id +
        '/field/' + PRIMARY_SERVICE_LINE_FIELD_ID;

      const clearResponse = UrlFetchApp.fetch(clearUrl, {
        method: 'delete',
        headers: { Authorization: token },
        muteHttpExceptions: true
      });

      const code = clearResponse.getResponseCode();

      if (code === 200) {
        clearedCount++;
        Logger.log('Cleared Primary Service Line on subtask "' + taskName + '"');
      } else {
        Logger.log('ERROR clearing "' + taskName + '" (status ' + code + '): ' + clearResponse.getContentText());
      }
    } catch (err) {
      Logger.log('ERROR clearing "' + taskName + '": ' + err);
    }
  });

  Logger.log(
    (DRY_RUN ? 'DRY RUN complete. Would clear ' : 'Done. Cleared ') +
    clearedCount + ' subtasks.'
  );
}
function debugListAllEmployeeFolders() {
  const FOLDER_GROUPS = {
    Infusion: [
      '1abALtYC_Bcdnl-J06wtOxyJQI6XFdpQs',
      '1qF4Nv_MLLOEbd4i8nb6PF_2TsnryH0o2',
      '1B4ZbEOPkQ8GW-RETHpWLcQDob137SRhG',
      '1Y3kbrNn_v2E8oyJQYCiH0Swj8m2bmwna'
    ],
    Vascular: ['185swROseZQ4U1nRhHtD-pFNIRx0DIz19'],
    Corporate: ['1UfIBI1wHcMbV_cz9KZSVfI8ud933gOBO']
  };

  const counts = { Infusion: 0, Vascular: 0, Corporate: 0 };

  Object.keys(FOLDER_GROUPS).forEach(function(serviceLine) {
    FOLDER_GROUPS[serviceLine].forEach(function(folderId) {
      try {
        const parentFolder = DriveApp.getFolderById(folderId);

        Logger.log('=== PARENT [' + serviceLine + ']: ' + parentFolder.getName() + ' (' + folderId + ') ===');

        const employeeFolders = parentFolder.getFolders();

        while (employeeFolders.hasNext()) {
          const employeeFolder = employeeFolders.next();
          Logger.log('  [' + serviceLine + '] ' + employeeFolder.getName());
          counts[serviceLine]++;
        }
      } catch (err) {
        Logger.log('ERROR opening folder ' + folderId + ' [' + serviceLine + ']: ' + err);
      }
    });
  });

  Logger.log(
    'TOTALS — Infusion: ' + counts.Infusion +
    ', Vascular: ' + counts.Vascular +
    ', Corporate: ' + counts.Corporate
  );
}
function fixCurrentTimecardDateAndBanner() {
  const WRONG_PAY_DATE = '05-22-2026';
  const CORRECT_PAY_DATE = '06-05-2026';
  const CORRECT_PERIOD_START_TEXT = '5/18';
  const CORRECT_PERIOD_END_TEXT = '5/31';
  const DRY_RUN = true;

  const PARENT_FOLDER_IDS = [
    '1abALtYC_Bcdnl-J06wtOxyJQI6XFdpQs',
    '1qF4Nv_MLLOEbd4i8nb6PF_2TsnryH0o2',
    '1B4ZbEOPkQ8GW-RETHpWLcQDob137SRhG',
    '1Y3kbrNn_v2E8oyJQYCiH0Swj8m2bmwna',
    '185swROseZQ4U1nRhHtD-pFNIRx0DIz19',
    '1UfIBI1wHcMbV_cz9KZSVfI8ud933gOBO'
  ];

  const bannerText =
    'Pay Period: ' + CORRECT_PERIOD_START_TEXT + ' - ' + CORRECT_PERIOD_END_TEXT;

  let renamedCount = 0;

  PARENT_FOLDER_IDS.forEach(function(parentFolderId) {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const employeeFolders = parentFolder.getFolders();

    while (employeeFolders.hasNext()) {
      const employeeFolder = employeeFolders.next();
      const files = employeeFolder.getFiles();

      while (files.hasNext()) {
        const file = files.next();
        const fileName = file.getName();

        if (!fileName.includes('Timecard')) continue;
        if (!fileName.includes(WRONG_PAY_DATE)) continue;

        const newName = fileName.replace(WRONG_PAY_DATE, CORRECT_PAY_DATE);

        try {
          if (DRY_RUN) {
            Logger.log('DRY RUN - UPDATED BANNER: ' + fileName);
            Logger.log('DRY RUN - RENAMED: ' + fileName + ' -> ' + newName);
            renamedCount++;
            continue;
          }

          const spreadsheet = SpreadsheetApp.openById(file.getId());
          const timecardSheet = spreadsheet.getSheetByName('TimeCardNotes');

          if (timecardSheet) {
            const bannerRange = timecardSheet.getRange('C1:E1');

            bannerRange.setValue(bannerText);
            bannerRange.setFontWeight('bold');
            bannerRange.setFontSize(14);
            bannerRange.setBackground('#FFF2CC');
            bannerRange.setHorizontalAlignment('center');
            bannerRange.setVerticalAlignment('middle');

            Logger.log('UPDATED BANNER: ' + fileName);
          } else {
            Logger.log('WARNING: No TimeCardNotes sheet in ' + fileName);
          }

          file.setName(newName);
          Logger.log('RENAMED: ' + fileName + ' -> ' + newName);

          renamedCount++;
        } catch (err) {
          Logger.log('ERROR processing ' + fileName + ': ' + err);
        }
      }
    }
  });

  Logger.log(
    (DRY_RUN ? 'DRY RUN complete. Would process ' : 'Done. Processed ') +
    renamedCount + ' timecard files.'
  );
}
function copyPriorEndingToCurrentBeginning() {
  const DRY_RUN = true;

  const CURRENT_PAY_DATE_STRING = '06-05-2026';

  const INFUSION_ARCHIVE_PERIOD_FOLDER_ID = '1xhwiI-IpEA2POzXftHN2NkGZrLDI2fxE';
  const VASCULAR_ARCHIVE_PERIOD_FOLDER_ID = '1c37xYxv9oozS4lGHB7PzEiVCf0bcdtrT';

  const INVENTORY_SHEET_NAME = 'Inventory';

  const SERVICE_LINES = [
    {
      name: 'Infusion',
      archivePeriodFolderId: INFUSION_ARCHIVE_PERIOD_FOLDER_ID,
      currentParentIds: [
        '1abALtYC_Bcdnl-J06wtOxyJQI6XFdpQs',
        '1qF4Nv_MLLOEbd4i8nb6PF_2TsnryH0o2',
        '1B4ZbEOPkQ8GW-RETHpWLcQDob137SRhG',
        '1Y3kbrNn_v2E8oyJQYCiH0Swj8m2bmwna'
      ],
      config: { itemColumn: 2, beginningColumn: 3, endingColumn: 5, firstItemRow: 3 }
    },
    {
      name: 'Vascular',
      archivePeriodFolderId: VASCULAR_ARCHIVE_PERIOD_FOLDER_ID,
      currentParentIds: ['185swROseZQ4U1nRhHtD-pFNIRx0DIz19'],
      config: { itemColumn: 1, beginningColumn: 2, endingColumn: 4, firstItemRow: 3 }
    }
  ];

  let archiveFilesProcessed = 0;
  let employeesMatched = 0;
  let itemsWritten = 0;          // from Ending
  let itemsWroteCorrected = 0;   // from Beginning (flag review)
  let itemsPreserved = 0;
  let itemsMissing = 0;
  let itemsNew = 0;
  let employeesSkipped = 0;

  SERVICE_LINES.forEach(function(sl) {
    const cfg = sl.config;

    // Open the archive period folder directly by its hardcoded ID
    let archivePeriodFolder;
    try {
      archivePeriodFolder = DriveApp.getFolderById(sl.archivePeriodFolderId);
    } catch (err) {
      Logger.log('ERROR: could not open ' + sl.name + ' archive period folder ' + sl.archivePeriodFolderId + ': ' + err + '. Skipping service line.');
      return;
    }

    // Iterate every file directly inside the archive folder (no subfolders)
    const archiveFiles = archivePeriodFolder.getFiles();

    while (archiveFiles.hasNext()) {
      const archiveFile = archiveFiles.next();
      const archiveFileName = archiveFile.getName();

      // Only archived timecard files; skip processed/summary/archive artifacts
      if (!archiveFileName.includes('Timecard')) continue;
      if (!/Pay Date \d{2}-\d{2}-\d{4}/.test(archiveFileName)) continue;
      if (archiveFileName.includes('Processed')) continue;
      if (archiveFileName.includes('Summary')) continue;
      if (archiveFileName.includes('Archive')) continue;
      if (archiveFileName.includes('ZZ_ARCHIVE_PAYROLL')) continue;

      try {
        archiveFilesProcessed++;

        // Employee folder name = everything before " - Timecard"
        const employeeName = archiveFileName.split(' - Timecard')[0].trim();

        // Read the archived Inventory tab
        const archiveSs = SpreadsheetApp.openById(archiveFile.getId());
        const archiveInventory = archiveSs.getSheetByName(INVENTORY_SHEET_NAME);

        if (!archiveInventory) {
          Logger.log('SKIP: no Inventory tab in archived timecard for ' + employeeName);
          employeesSkipped++;
          continue;
        }

        const archiveNumRows = archiveInventory.getMaxRows() - cfg.firstItemRow + 1;
        const archiveItemNames = archiveInventory.getRange(cfg.firstItemRow, cfg.itemColumn, archiveNumRows, 1).getValues();
        const archiveBeginningVals = archiveInventory.getRange(cfg.firstItemRow, cfg.beginningColumn, archiveNumRows, 1).getValues();
        const archiveEndingVals = archiveInventory.getRange(cfg.firstItemRow, cfg.endingColumn, archiveNumRows, 1).getValues();

        const archiveItems = [];
        for (let i = 0; i < archiveNumRows; i++) {
          const itemName = archiveItemNames[i][0];
          if (itemName === '' || itemName === null) break;

          const beginningVal = archiveBeginningVals[i][0];
          const endingVal = archiveEndingVals[i][0];

          let trueValue = null;
          let source = null;

          if (endingVal !== '' && endingVal !== null) {
            trueValue = endingVal;
            source = 'Ending';
          } else if (beginningVal !== '' && beginningVal !== null) {
            trueValue = beginningVal;
            source = 'Beginning';
          }

          archiveItems.push({
            name: String(itemName),
            normalized: String(itemName).trim().toLowerCase(),
            trueValue: trueValue,
            source: source
          });
        }

        // Find the matching current employee folder (exact name match)
        let currentFolder = null;
        for (let p = 0; p < sl.currentParentIds.length && !currentFolder; p++) {
          const currentParent = DriveApp.getFolderById(sl.currentParentIds[p]);
          const folderMatches = currentParent.getFoldersByName(employeeName);
          if (folderMatches.hasNext()) {
            currentFolder = folderMatches.next();
          }
        }

        if (!currentFolder) {
          Logger.log('SKIP: no current folder for ' + employeeName);
          employeesSkipped++;
          continue;
        }

        // Find the current timecard file
        const currentFiles = currentFolder.getFiles();
        let currentTimecard = null;

        while (currentFiles.hasNext()) {
          const f = currentFiles.next();
          const fn = f.getName();
          if (fn.includes('Timecard') && fn.includes(CURRENT_PAY_DATE_STRING)) {
            currentTimecard = f;
            break;
          }
        }

        if (!currentTimecard) {
          Logger.log('SKIP: no current timecard for ' + employeeName);
          employeesSkipped++;
          continue;
        }

        // Read the current Inventory tab and build item name -> row map
        const currentSs = SpreadsheetApp.openById(currentTimecard.getId());
        const currentInventory = currentSs.getSheetByName(INVENTORY_SHEET_NAME);

        if (!currentInventory) {
          Logger.log('SKIP: no Inventory tab in current timecard for ' + employeeName);
          employeesSkipped++;
          continue;
        }

        employeesMatched++;

        const currentNumRows = currentInventory.getMaxRows() - cfg.firstItemRow + 1;
        const currentItemNames = currentInventory.getRange(cfg.firstItemRow, cfg.itemColumn, currentNumRows, 1).getValues();
        const currentBeginningVals = currentInventory.getRange(cfg.firstItemRow, cfg.beginningColumn, currentNumRows, 1).getValues();

        const currentItems = [];
        const currentRowByName = {};

        for (let i = 0; i < currentNumRows; i++) {
          const itemName = currentItemNames[i][0];
          if (itemName === '' || itemName === null) break;

          const normalized = String(itemName).trim().toLowerCase();
          const row = cfg.firstItemRow + i;

          currentItems.push({ name: String(itemName), normalized: normalized });
          currentRowByName[normalized] = { row: row, beginning: currentBeginningVals[i][0] };
        }

        const matchedCurrent = {};

        // Carry each archived item (that has a true value) into current Beginning
        archiveItems.forEach(function(ai) {
          if (ai.trueValue === null) return; // archive had no data for this item

          const match = currentRowByName[ai.normalized];

          if (!match) {
            Logger.log('MISSING ITEM: ' + employeeName + ' - ' + ai.name + ' not found in current timecard');
            itemsMissing++;
            return;
          }

          matchedCurrent[ai.normalized] = true;

          const currentValue = match.beginning;
          const isBlank = (currentValue === '' || currentValue === null);

          if (!isBlank) {
            if (!DRY_RUN) {
              currentInventory.getRange(match.row, cfg.beginningColumn)
                .setHorizontalAlignment('center')
                .setVerticalAlignment('middle');
            }
            Logger.log((DRY_RUN ? 'DRY RUN - ' : '') + 'PRESERVED: ' + employeeName + ' - ' + ai.name + ' - existing value ' + currentValue + ', would have written ' + ai.trueValue + ' (source: ' + ai.source + ')');
            itemsPreserved++;
            return;
          }

          if (!DRY_RUN) {
            currentInventory.getRange(match.row, cfg.beginningColumn)
              .setValue(ai.trueValue)
              .setHorizontalAlignment('center')
              .setVerticalAlignment('middle');
          }

          if (ai.source === 'Ending') {
            Logger.log((DRY_RUN ? 'DRY RUN - ' : '') + 'WROTE: ' + employeeName + ' - ' + ai.name + ' -> ' + ai.trueValue + ' (from Ending)');
            itemsWritten++;
          } else {
            Logger.log((DRY_RUN ? 'DRY RUN - ' : '') + 'WROTE-CORRECTED: ' + employeeName + ' - ' + ai.name + ' -> ' + ai.trueValue + ' (from archive Beginning — flag review)');
            itemsWroteCorrected++;
          }
        });

        // Items present in current but not in archive
        currentItems.forEach(function(ci) {
          if (!matchedCurrent[ci.normalized]) {
            Logger.log('NEW ITEM: ' + employeeName + ' - ' + ci.name + ' in current but not in archive');
            itemsNew++;
          }
        });
      } catch (err) {
        Logger.log('ERROR processing archive file "' + archiveFileName + '" [' + sl.name + ']: ' + err);
      }
    }
  });

  Logger.log(
    (DRY_RUN ? 'DRY RUN complete. ' : 'Done. ') +
    'Archive files processed: ' + archiveFilesProcessed +
    ', Employees matched: ' + employeesMatched +
    ', Items written (from Ending): ' + itemsWritten +
    ', Items wrote-corrected (from Beginning — flag review): ' + itemsWroteCorrected +
    ', Items preserved: ' + itemsPreserved +
    ', Items missing: ' + itemsMissing +
    ', Items new: ' + itemsNew +
    ', Employees skipped: ' + employeesSkipped
  );
}
function debugInspectArchivePeriodFolders() {
  const ARCHIVE_PERIOD_FOLDER_IDS = {
    Infusion: '1xhwiI-IpEA2POzXftHN2NkGZrLDI2fxE',
    Vascular: '1c37xYxv9oozS4lGHB7PzEiVCf0bcdtrT'
  };

  Object.keys(ARCHIVE_PERIOD_FOLDER_IDS).forEach(function(serviceLine) {
    const folderId = ARCHIVE_PERIOD_FOLDER_IDS[serviceLine];

    try {
      const folder = DriveApp.getFolderById(folderId);
      Logger.log('=== ' + serviceLine + ' archive period folder: ' + folder.getName() + ' (' + folderId + ') ===');

      // Direct subfolders
      const subfolders = folder.getFolders();
      let subfolderCount = 0;
      while (subfolders.hasNext()) {
        const sf = subfolders.next();
        subfolderCount++;
        if (subfolderCount <= 10) {
          Logger.log('  SUBFOLDER: ' + sf.getName());
        }
      }
      Logger.log('  Total subfolders directly inside: ' + subfolderCount);

      // Direct files
      const files = folder.getFiles();
      let fileCount = 0;
      while (files.hasNext()) {
        const f = files.next();
        fileCount++;
        if (fileCount <= 10) {
          Logger.log('  FILE: ' + f.getName());
        }
      }
      Logger.log('  Total files directly inside: ' + fileCount);
    } catch (err) {
      Logger.log('ERROR opening ' + serviceLine + ' archive period folder ' + folderId + ': ' + err);
    }
  });
}
function debugArchiveFolderStructure() {
  const INFUSION_ARCHIVE_PERIOD_FOLDER_ID = '1xhwiI-IpEA2POzXftHN2NkGZrLDI2fxE';
  const VASCULAR_ARCHIVE_PERIOD_FOLDER_ID = '1c37xYxv9oozS4lGHB7PzEiVCf0bcdtrT';

  const ARCHIVES = [
    { name: 'Infusion', id: INFUSION_ARCHIVE_PERIOD_FOLDER_ID },
    { name: 'Vascular', id: VASCULAR_ARCHIVE_PERIOD_FOLDER_ID }
  ];

  ARCHIVES.forEach(function(archive) {
    try {
      const folder = DriveApp.getFolderById(archive.id);

      Logger.log('=== ARCHIVE [' + archive.name + ']: ' + folder.getName() + ' ===');

      // Subfolders (log all names, peek inside the first 3)
      const subfolders = folder.getFolders();
      let subfolderCount = 0;

      while (subfolders.hasNext()) {
        const subfolder = subfolders.next();
        subfolderCount++;

        Logger.log('  SUBFOLDER: ' + subfolder.getName());

        if (subfolderCount <= 3) {
          const subFiles = subfolder.getFiles();
          while (subFiles.hasNext()) {
            const sf = subFiles.next();
            Logger.log('    FILE in ' + subfolder.getName() + ': ' + sf.getName());
          }
        }
      }

      // Files directly inside the archive folder
      const files = folder.getFiles();
      let fileCount = 0;

      while (files.hasNext()) {
        const file = files.next();
        fileCount++;
        Logger.log('  FILE: ' + file.getName());
      }

      Logger.log('  TOTALS [' + archive.name + '] — subfolders: ' + subfolderCount + ', files: ' + fileCount);
    } catch (err) {
      Logger.log('ERROR opening ' + archive.name + ' archive folder ' + archive.id + ': ' + err);
    }
  });
}
function debugArchiveCurrentFolderMatches() {
  const INFUSION_ARCHIVE_PERIOD_FOLDER_ID = '1xhwiI-IpEA2POzXftHN2NkGZrLDI2fxE';
  const VASCULAR_ARCHIVE_PERIOD_FOLDER_ID = '1c37xYxv9oozS4lGHB7PzEiVCf0bcdtrT';

  const SERVICE_LINES = [
    {
      name: 'Infusion',
      archivePeriodFolderId: INFUSION_ARCHIVE_PERIOD_FOLDER_ID,
      currentParentIds: [
        '1abALtYC_Bcdnl-J06wtOxyJQI6XFdpQs',
        '1qF4Nv_MLLOEbd4i8nb6PF_2TsnryH0o2',
        '1B4ZbEOPkQ8GW-RETHpWLcQDob137SRhG',
        '1Y3kbrNn_v2E8oyJQYCiH0Swj8m2bmwna'
      ]
    },
    {
      name: 'Vascular',
      archivePeriodFolderId: VASCULAR_ARCHIVE_PERIOD_FOLDER_ID,
      currentParentIds: ['185swROseZQ4U1nRhHtD-pFNIRx0DIz19']
    }
  ];

  SERVICE_LINES.forEach(function(sl) {
    // Build the list of current employee folder names
    const currentFolderNames = [];

    sl.currentParentIds.forEach(function(parentId) {
      try {
        const parentFolder = DriveApp.getFolderById(parentId);
        const subfolders = parentFolder.getFolders();
        while (subfolders.hasNext()) {
          currentFolderNames.push(subfolders.next().getName());
        }
      } catch (err) {
        Logger.log('ERROR opening current parent ' + parentId + ' [' + sl.name + ']: ' + err);
      }
    });

    Logger.log('=== CURRENT [' + sl.name + '] FOLDERS ===');
    currentFolderNames.forEach(function(n) {
      Logger.log('  ' + n);
    });

    // Exact-name lookup (mirrors getFoldersByName used by the real sync)
    const currentSet = {};
    currentFolderNames.forEach(function(n) {
      currentSet[n] = true;
    });

    let matches = 0;
    let nonMatches = 0;

    let archivePeriodFolder;
    try {
      archivePeriodFolder = DriveApp.getFolderById(sl.archivePeriodFolderId);
    } catch (err) {
      Logger.log('ERROR opening ' + sl.name + ' archive period folder ' + sl.archivePeriodFolderId + ': ' + err);
      return;
    }

    Logger.log('=== ARCHIVE MATCH RESULTS [' + sl.name + '] ===');

    const archiveFiles = archivePeriodFolder.getFiles();

    while (archiveFiles.hasNext()) {
      const archiveFile = archiveFiles.next();
      const archiveFileName = archiveFile.getName();

      if (!archiveFileName.includes('Timecard')) continue;
      if (!/Pay Date \d{2}-\d{2}-\d{4}/.test(archiveFileName)) continue;
      if (archiveFileName.includes('Processed')) continue;
      if (archiveFileName.includes('Summary')) continue;
      if (archiveFileName.includes('ZZ_ARCHIVE_PAYROLL')) continue;

      const archiveName = archiveFileName.split(' - Timecard')[0].trim();

      if (currentSet[archiveName]) {
        Logger.log('MATCH: ' + archiveName);
        matches++;
      } else {
        // Suggest near-misses: current folders starting with the same last-name prefix
        const lastNamePart = (archiveName.split(',')[0] || '').trim().toLowerCase();
        const prefix = lastNamePart.substring(0, 4);

        const closest = currentFolderNames.filter(function(n) {
          return prefix.length > 0 && n.trim().toLowerCase().indexOf(prefix) === 0;
        });

        Logger.log('NO MATCH: ' + archiveName + ' — closest current folders containing the last name: ' + (closest.length > 0 ? closest.join(', ') : 'none'));
        nonMatches++;
      }
    }

    Logger.log('TOTALS [' + sl.name + '] — matches: ' + matches + ', non-matches: ' + nonMatches);
  });
}

function buildMasterInventorySummary() {
  const MASTER_SPREADSHEET_ID = '1QQlpd-gwe07yWOTGpmixzv9ndtvkQFC2alPzyi5jvq4';
  const INFUSION_TAB_NAME = 'Infusion';
  const VASCULAR_TAB_NAME = 'Vascular';
  const INFUSION_ARCHIVE_PERIOD_FOLDER_ID = '1xhwiI-IpEA2POzXftHN2NkGZrLDI2fxE';
  const VASCULAR_ARCHIVE_PERIOD_FOLDER_ID = '1c37xYxv9oozS4lGHB7PzEiVCf0bcdtrT';

  const INVENTORY_SHEET_NAME = 'Inventory';
  const DRY_RUN = true;

  const HEADER = [
    'Employee', 'Item', 'Beginning Count', 'Ending Count',
    'Used During Period', 'Total Value', 'Status', 'Corrections / Notes'
  ];

  const SERVICE_LINES = [
    {
      name: 'Infusion',
      tabName: INFUSION_TAB_NAME,
      archivePeriodFolderId: INFUSION_ARCHIVE_PERIOD_FOLDER_ID,
      currentParentIds: [
        '1abALtYC_Bcdnl-J06wtOxyJQI6XFdpQs',
        '1qF4Nv_MLLOEbd4i8nb6PF_2TsnryH0o2',
        '1B4ZbEOPkQ8GW-RETHpWLcQDob137SRhG',
        '1Y3kbrNn_v2E8oyJQYCiH0Swj8m2bmwna'
      ],
      config: { itemColumn: 2, beginningColumn: 3, endingColumn: 5, usedColumn: 6, totalValueColumn: 9, firstItemRow: 3 }
    },
    {
      name: 'Vascular',
      tabName: VASCULAR_TAB_NAME,
      archivePeriodFolderId: VASCULAR_ARCHIVE_PERIOD_FOLDER_ID,
      currentParentIds: ['185swROseZQ4U1nRhHtD-pFNIRx0DIz19'],
      config: { itemColumn: 1, beginningColumn: 2, endingColumn: 4, usedColumn: 5, totalValueColumn: 8, firstItemRow: 3 }
    }
  ];

  function normalizeName(name) {
    let n = String(name).toLowerCase();
    n = n.replace(/[_ ](ft|prn|pt|1099)$/i, '');
    n = n.replace(/^\d+\.\s*/, '');
    n = n.replace(/\([^)]*\)/g, ' ');
    n = n.replace(/[,._()]/g, ' ');
    n = n.replace(/\s+/g, ' ').trim();

    let prev;
    do {
      prev = n;
      n = n.replace(/\s+(ft|prn|pt|1099)$/i, '').trim();
    } while (n !== prev);

    n = n.split(/\s+/).sort().join(' ');
    return n;
  }

  function isBlankValue(v) {
    return v === '' || v === null;
  }

  let master = null;
  if (!DRY_RUN) {
    master = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  }

  SERVICE_LINES.forEach(function(sl) {
    const cfg = sl.config;

    // Build current employee folder lookup (normalized name -> display name)
    const currentLookup = {};
    sl.currentParentIds.forEach(function(parentId) {
      try {
        const parentFolder = DriveApp.getFolderById(parentId);
        const subfolders = parentFolder.getFolders();
        while (subfolders.hasNext()) {
          const sf = subfolders.next();
          const norm = normalizeName(sf.getName());
          if (!norm || norm.includes('archive')) continue;
          currentLookup[norm] = { name: sf.getName() };
        }
      } catch (err) {
        Logger.log('ERROR opening current parent ' + parentId + ' [' + sl.name + ']: ' + err);
      }
    });

    const rows = [];
    const archiveNormsSeen = {};

    // Iterate archive timecard files
    let archivePeriodFolder;
    try {
      archivePeriodFolder = DriveApp.getFolderById(sl.archivePeriodFolderId);
    } catch (err) {
      Logger.log('ERROR opening ' + sl.name + ' archive period folder ' + sl.archivePeriodFolderId + ': ' + err);
      return;
    }

    const archiveFiles = archivePeriodFolder.getFiles();

    while (archiveFiles.hasNext()) {
      const archiveFile = archiveFiles.next();
      const archiveFileName = archiveFile.getName();

      if (!archiveFileName.includes('Timecard')) continue;
      if (archiveFileName.includes('Processed')) continue;
      if (archiveFileName.includes('Summary')) continue;
      if (archiveFileName.includes('ZZ_ARCHIVE_PAYROLL')) continue;

      try {
        const employeeName = archiveFileName.split(' - Timecard')[0].trim();
        const archiveNorm = normalizeName(employeeName);
        archiveNormsSeen[archiveNorm] = true;

        const archiveSs = SpreadsheetApp.openById(archiveFile.getId());
        const inv = archiveSs.getSheetByName(INVENTORY_SHEET_NAME);

        if (!inv) {
          Logger.log('SKIP: no Inventory tab in archived timecard for ' + employeeName);
          continue;
        }

        const numRows = inv.getMaxRows() - cfg.firstItemRow + 1;
        const itemVals = inv.getRange(cfg.firstItemRow, cfg.itemColumn, numRows, 1).getValues();
        const beginVals = inv.getRange(cfg.firstItemRow, cfg.beginningColumn, numRows, 1).getValues();
        const endVals = inv.getRange(cfg.firstItemRow, cfg.endingColumn, numRows, 1).getValues();
        const usedVals = inv.getRange(cfg.firstItemRow, cfg.usedColumn, numRows, 1).getValues();
        const totalVals = inv.getRange(cfg.firstItemRow, cfg.totalValueColumn, numRows, 1).getValues();

        const employeeRows = [];

        for (let i = 0; i < numRows; i++) {
          const itemName = itemVals[i][0];
          if (isBlankValue(itemName)) break;

          const beginning = beginVals[i][0];
          const ending = endVals[i][0];
          const used = usedVals[i][0];
          const totalValue = totalVals[i][0];

          let status;
          let note;

          if (!isBlankValue(ending)) {
            status = 'Entered correctly';
            note = '';
          } else if (!isBlankValue(beginning)) {
            status = 'Wrong column — entered in Beginning';
            note = 'Please review — entered in Beginning column instead of Ending';
          } else {
            status = 'Missing';
            note = 'Please review — no count entered';
          }

          employeeRows.push([
            employeeName, String(itemName), beginning, ending, used, totalValue, status, note
          ]);
        }

        // Override status if this archived employee has no current folder
        if (!currentLookup[archiveNorm]) {
          employeeRows.forEach(function(r) {
            r[6] = 'Termed/not found in current';
            r[7] = 'Verify employment status';
          });
        }

        employeeRows.forEach(function(r) {
          rows.push(r);
        });
      } catch (err) {
        Logger.log('ERROR processing archive file "' + archiveFileName + '" [' + sl.name + ']: ' + err);
      }
    }

    // Current employees with no archive data
    Object.keys(currentLookup).forEach(function(norm) {
      if (!archiveNormsSeen[norm]) {
        rows.push([
          currentLookup[norm].name, '(no items found)', '', '', '', '',
          'No archive data — manual entry needed', 'Verify whether nurse had prior period activity.'
        ]);
      }
    });

    // Sort by Employee (A) then Item (B)
    rows.sort(function(a, b) {
      const ea = String(a[0]).toLowerCase();
      const eb = String(b[0]).toLowerCase();
      if (ea < eb) return -1;
      if (ea > eb) return 1;

      const ia = String(a[1]).toLowerCase();
      const ib = String(b[1]).toLowerCase();
      if (ia < ib) return -1;
      if (ia > ib) return 1;
      return 0;
    });

    // Status counts
    const statusCounts = {};
    rows.forEach(function(r) {
      const s = r[6];
      statusCounts[s] = (statusCounts[s] || 0) + 1;
    });

    Logger.log('=== ' + sl.name + ' tab — ' + rows.length + ' data rows ===');

    if (DRY_RUN) {
      Logger.log('DRY RUN - would write ' + rows.length + ' rows to ' + sl.tabName);
      rows.slice(0, 5).forEach(function(r) {
        Logger.log('  SAMPLE: ' + r.join(' | '));
      });
    } else {
      let tab = master.getSheetByName(sl.tabName);
      if (!tab) {
        tab = master.insertSheet(sl.tabName);
      }

      tab.clear();
      tab.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight('bold');
      tab.setFrozenRows(1);

      if (rows.length > 0) {
        tab.getRange(2, 1, rows.length, HEADER.length).setValues(rows);
      }

      tab.autoResizeColumn(1);
      tab.autoResizeColumn(2);
    }

    Object.keys(statusCounts).forEach(function(s) {
      Logger.log('  STATUS [' + sl.name + '] ' + s + ': ' + statusCounts[s]);
    });
  });

  Logger.log(DRY_RUN ? 'DRY RUN complete. No changes written to master sheet.' : 'Done. Master inventory summary rebuilt.');
}
function testDateLogic() {
  function fmt(d) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MM-dd-yyyy');
  }

  const cases = [
    {
      runDate: new Date(2026, 5, 3),   // June 3, 2026 (next scheduled run)
      expectedStart: '06-01-2026',
      expectedEnd: '06-14-2026',
      expectedPay: '06-19-2026'
    },
    {
      runDate: new Date(2026, 5, 17),  // June 17, 2026
      expectedStart: '06-15-2026',
      expectedEnd: '06-28-2026',
      expectedPay: '07-03-2026'
    },
    {
      runDate: new Date(2026, 6, 1),   // July 1, 2026
      expectedStart: '06-29-2026',
      expectedEnd: '07-12-2026',
      expectedPay: '07-17-2026'
    }
  ];

  cases.forEach(function(c) {
    const info = getCurrentRunInfo(c.runDate);

    const gotStart = fmt(info.currentPeriodStart);
    const gotEnd = fmt(info.currentPeriodEnd);
    const gotPay = fmt(info.payDate);

    const pass =
      gotStart === c.expectedStart &&
      gotEnd === c.expectedEnd &&
      gotPay === c.expectedPay;

    Logger.log(
      (pass ? 'PASS' : 'FAIL') + ' — Run ' + fmt(c.runDate) +
      ' -> period ' + gotStart + ' to ' + gotEnd + ', pay date ' + gotPay +
      (pass ? '' : ' (expected period ' + c.expectedStart + ' to ' + c.expectedEnd + ', pay date ' + c.expectedPay + ')')
    );
  });
}
function replaceInventoryTabsWithFormLink() {
  const CURRENT_PAY_DATE_STRING = '06-05-2026';
  const PAY_PERIOD_DATES = '5/18 - 5/31';
  const VASCULAR_FORM_URL = 'https://forms.clickup.com/9017962545/f/8cr6c1h-8417/S1B47HI0Z0GR2VM7L2';
  const VASCULAR_PARENT_FOLDER_ID = '185swROseZQ4U1nRhHtD-pFNIRx0DIz19';
  const INVENTORY_SHEET_NAME = 'Inventory';
  const DRY_RUN = true;

  let foldersWalked = 0;
  let updated = 0;
  let skipped = 0;
  let errors = 0;

  const parentFolder = DriveApp.getFolderById(VASCULAR_PARENT_FOLDER_ID);
  const employeeFolders = parentFolder.getFolders();

  while (employeeFolders.hasNext()) {
    const employeeFolder = employeeFolders.next();
    const employeeName = employeeFolder.getName();

    if (employeeName.toLowerCase().includes('archive')) continue;

    foldersWalked++;

    // Find the current timecard spreadsheet
    const files = employeeFolder.getFiles();
    let timecard = null;

    while (files.hasNext()) {
      const f = files.next();
      const fn = f.getName();

      if (!fn.includes('Timecard')) continue;
      if (!fn.includes(CURRENT_PAY_DATE_STRING)) continue;
      if (fn.includes('Processed')) continue;
      if (fn.includes('Summary')) continue;
      if (fn.includes('ZZ_ARCHIVE_PAYROLL')) continue;

      timecard = f;
      break;
    }

    if (!timecard) {
      Logger.log('SKIPPED (no timecard found): ' + employeeName);
      skipped++;
      continue;
    }

    try {
      const ss = SpreadsheetApp.openById(timecard.getId());
      const inv = ss.getSheetByName(INVENTORY_SHEET_NAME);

      if (!inv) {
        Logger.log('SKIPPED (no inventory tab): ' + employeeName);
        skipped++;
        continue;
      }

      if (DRY_RUN) {
        Logger.log('DRY RUN - WROTE: ' + employeeName + ' (' + timecard.getName() + ')');
        updated++;
        continue;
      }

      // Wipe the tab entirely (values, formulas, formatting, merges)
      inv.getRange(1, 1, inv.getMaxRows(), inv.getMaxColumns()).breakApart();
      inv.clear();

      // Header banner across A1:E1
      inv.getRange('A1:E1').merge();
      const headerCell = inv.getRange('A1');
      headerCell.setValue('Inventory — Pay Period: ' + PAY_PERIOD_DATES);
      headerCell.setFontWeight('bold');
      headerCell.setFontSize(14);
      headerCell.setBackground('#FFF2CC');
      headerCell.setHorizontalAlignment('center');

      // Inventory form hyperlink across A3:E3
      inv.getRange('A3:E3').merge();
      const linkCell = inv.getRange('A3');
      linkCell.setFormula('=HYPERLINK("' + VASCULAR_FORM_URL + '", "ENTER YOUR INVENTORY HERE")');
      linkCell.setFontWeight('bold');
      linkCell.setFontSize(18);
      linkCell.setHorizontalAlignment('center');
      inv.setRowHeight(3, 50);

      // Auto-resize columns A through E
      for (let c = 1; c <= 5; c++) {
        inv.autoResizeColumn(c);
      }

      Logger.log('WROTE: ' + employeeName + ' (' + timecard.getName() + ')');
      updated++;
    } catch (err) {
      Logger.log('ERROR: ' + employeeName + ': ' + err);
      errors++;
    }
  }

  Logger.log(
    (DRY_RUN ? 'DRY RUN complete. ' : 'Done. ') +
    'Folders walked: ' + foldersWalked +
    ', Timecards updated: ' + updated +
    ', Skipped: ' + skipped +
    ', Errors: ' + errors
  );
}
function convertExcelToGoogleSheetsInVascularFolders() {
  const VASCULAR_PARENT_FOLDER_ID = '185swROseZQ4U1nRhHtD-pFNIRx0DIz19';
  const DRY_RUN = true;

  let foldersWalked = 0;
  let converted = 0;
  let skipped = 0;
  let errors = 0;

  const parentFolder = DriveApp.getFolderById(VASCULAR_PARENT_FOLDER_ID);
  const employeeFolders = parentFolder.getFolders();

  while (employeeFolders.hasNext()) {
    const employeeFolder = employeeFolders.next();
    const employeeName = employeeFolder.getName();

    if (employeeName.toLowerCase().includes('archive')) continue;

    foldersWalked++;

    const files = employeeFolder.getFiles();

    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();

      if (!fileName.toLowerCase().endsWith('.xlsx')) continue;

      const newName = fileName.replace(/\.xlsx$/i, '');

      try {
        // Skip if a Google Sheet with the target name already exists in the folder
        const existing = employeeFolder.getFilesByName(newName);
        if (existing.hasNext()) {
          Logger.log('SKIPPED (already converted): ' + employeeName + ' / ' + fileName);
          skipped++;
          continue;
        }

        if (DRY_RUN) {
          Logger.log('DRY RUN - WOULD CONVERT + TRASH: ' + employeeName + ' / ' + fileName + ' -> ' + newName);
          converted++;
          continue;
        }

        const convertedFile = Drive.Files.copy(
          { title: newName, mimeType: MimeType.GOOGLE_SHEETS },
          file.getId()
        );

        const newGoogleSheet = DriveApp.getFileById(convertedFile.id);
        employeeFolder.addFile(newGoogleSheet);

        file.setTrashed(true);

        Logger.log('CONVERTED + TRASHED: ' + employeeName + ' / ' + fileName + ' -> ' + newName);
        converted++;
      } catch (err) {
        Logger.log('ERROR converting ' + employeeName + ' / ' + fileName + ': ' + err);
        errors++;
      }
    }
  }

  Logger.log(
    (DRY_RUN ? 'DRY RUN complete. ' : 'Done. ') +
    'Folders walked: ' + foldersWalked +
    ', Converted' + (DRY_RUN ? ' (would)' : '') + ': ' + converted +
    ', Skipped: ' + skipped +
    ', Errors: ' + errors
  );
}
function trashAllSheetsInVascularFolders() {
  const VASCULAR_PARENT_FOLDER_ID = '185swROseZQ4U1nRhHtD-pFNIRx0DIz19';
  const DRY_RUN = true;

  let foldersWalked = 0;
  let trashed = 0;
  let errors = 0;

  const parentFolder = DriveApp.getFolderById(VASCULAR_PARENT_FOLDER_ID);
  const employeeFolders = parentFolder.getFolders();

  while (employeeFolders.hasNext()) {
    const employeeFolder = employeeFolders.next();
    const employeeName = employeeFolder.getName();

    if (employeeName.toLowerCase().includes('archive')) continue;

    foldersWalked++;

    const files = employeeFolder.getFiles();

    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      const mimeType = file.getMimeType();

      const isSheet =
        mimeType === MimeType.GOOGLE_SHEETS ||
        mimeType === MimeType.MICROSOFT_EXCEL ||
        fileName.toLowerCase().endsWith('.xlsx');

      if (!isSheet) continue;

      try {
        if (DRY_RUN) {
          Logger.log('DRY RUN - WOULD TRASH: ' + employeeName + ' / ' + fileName);
        } else {
          file.setTrashed(true);
          Logger.log('TRASHED: ' + employeeName + ' / ' + fileName);
        }
        trashed++;
      } catch (err) {
        Logger.log('ERROR trashing ' + employeeName + ' / ' + fileName + ': ' + err);
        errors++;
      }
    }
  }

  Logger.log(
    (DRY_RUN ? 'DRY RUN complete. ' : 'Done. ') +
    'Folders walked: ' + foldersWalked +
    ', Sheets ' + (DRY_RUN ? 'that would be trashed' : 'trashed') + ': ' + trashed +
    ', Errors: ' + errors
  );
}
