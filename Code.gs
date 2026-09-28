/*******************************************************
 * TTB LABEL COMPLIANCE ASSISTANT
 * Google Apps Script Backend
 *
 * PROTOTYPE — NOT AN OFFICIAL TTB SYSTEM
 *******************************************************/

const SPREADSHEET_ID =
  '1cO48ia6NbX-bkj5DXeid5FkRPf7lbSZ_e9a7_ZDoSAE';

const WEB_APP_URL =
  'https://script.google.com/macros/s/AKfycbyhpwOACzTCF1cZde48bx1sglp6Tl9UqUYfXCYpXkkruYcT1KOHix1DNtclHUE4WNSc/exec';

const SHEETS = {
  LABELS: 'Label Applications',
  PROFILES: 'Class Type Profiles'
};

/*******************************************************
 * STANDARDIZED DATABASE COLUMNS
 *******************************************************/

const LABEL_HEADERS = [
  'Record ID',
  'Timestamp',
  'Beverage Category',
  'Brand Name',
  'Class/Type Designation',
  'Alcohol Content',
  'Alcohol Content Required',
  'Net Contents',
  'Bottler Name',
  'Bottler Address',
  'Producer Name',
  'Producer Address',
  'Same Producer/Bottler Address',
  'Imported',
  'Country of Origin',
  'Government Health Warning Statement Provided',
  'Warning Exact Match',
  'Class/Type Profile ID',
  'Review Status'
];

const PROFILE_HEADERS = [
  'Profile ID',
  'Class/Type Designation',
  'Beverage Category',
  'Submission Count',
  'First Seen',
  'Last Seen',
  'Last Review Status'
];

/*******************************************************
 * REQUIRED GOVERNMENT HEALTH WARNING
 *******************************************************/

const WARNING_TEXT =
  'GOVERNMENT WARNING: (1) According to the Surgeon General, ' +
  'women should not drink alcoholic beverages during pregnancy ' +
  'because of the risk of birth defects. (2) Consumption of ' +
  'alcoholic beverages impairs your ability to drive a car or ' +
  'operate machinery, and may cause health problems.';

/*******************************************************
 * WEB APP
 *******************************************************/

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('TTB Label Compliance Assistant')
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );
}

/*******************************************************
 * INITIAL DATABASE SETUP
 *
 * Run this once if creating the database from scratch.
 *******************************************************/

function setupDatabase() {
  const ss = getSpreadsheet_();

  const labelSheet = ensureSheet_(
    ss,
    SHEETS.LABELS,
    LABEL_HEADERS
  );

  const profileSheet = ensureSheet_(
    ss,
    SHEETS.PROFILES,
    PROFILE_HEADERS
  );

  formatHeader_(labelSheet);
  formatHeader_(profileSheet);

  return {
    success: true,
    message: 'TTB prototype database initialized.',
    spreadsheetId: SPREADSHEET_ID
  };
}

/*******************************************************
 * IMPORTANT:
 * MIGRATE EXISTING GOOGLE SHEETS COLUMNS
 *
 * This function:
 * - Renames old columns
 * - Reorders columns
 * - Preserves existing data
 * - Removes obsolete columns
 * - Creates missing columns
 *
 * Run:
 * Apps Script → select alignPrototypeColumns → Run
 *******************************************************/

function alignPrototypeColumns() {
  const ss = getSpreadsheet_();

  const labelSheet = ensureSheet_(
    ss,
    SHEETS.LABELS,
    LABEL_HEADERS
  );

  const profileSheet = ensureSheet_(
    ss,
    SHEETS.PROFILES,
    PROFILE_HEADERS
  );

  /*****************************************************
   * LABEL APPLICATIONS
   *****************************************************/

  migrateSheetColumns_(
    labelSheet,
    LABEL_HEADERS,
    {
      'Government Warning':
        'Government Health Warning Statement Provided',

      'Government Health Warning':
        'Government Health Warning Statement Provided',

      'Health Warning':
        'Government Health Warning Statement Provided'
    }
  );

  /*****************************************************
   * CLASS TYPE PROFILES
   *****************************************************/

  migrateSheetColumns_(
    profileSheet,
    PROFILE_HEADERS,
    {}
  );

  formatHeader_(labelSheet);
  formatHeader_(profileSheet);

  return {
    success: true,
    message:
      'Google Sheets columns have been aligned with the TTB prototype.',
    labelColumns: LABEL_HEADERS,
    profileColumns: PROFILE_HEADERS
  };
}

/*******************************************************
 * COLUMN MIGRATION ENGINE
 *******************************************************/

function migrateSheetColumns_(
  sheet,
  desiredHeaders,
  renameMap
) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  /*****************************************************
   * EMPTY SHEET
   *****************************************************/

  if (lastRow === 0 || lastColumn === 0) {
    sheet
      .getRange(1, 1, 1, desiredHeaders.length)
      .setValues([desiredHeaders]);

    formatHeader_(sheet);
    return;
  }

  /*****************************************************
   * READ EXISTING DATA
   *****************************************************/

  const existingData = sheet
    .getRange(
      1,
      1,
      lastRow,
      lastColumn
    )
    .getValues();

  const existingHeaders = existingData[0].map(function(header) {
    const headerText = String(header || '').trim();

    return renameMap[headerText] || headerText;
  });

  /*****************************************************
   * MAP EXISTING COLUMNS BY HEADER
   *****************************************************/

  const headerIndex = {};

  existingHeaders.forEach(function(header, index) {
    if (header) {
      headerIndex[header] = index;
    }
  });

  /*****************************************************
   * BUILD NEW DATA IN REQUIRED ORDER
   *****************************************************/

  const newData = [];

  // Header row
  newData.push(desiredHeaders);

  // Existing data rows
  for (let row = 1; row < existingData.length; row++) {
    const newRow = [];

    desiredHeaders.forEach(function(header) {
      if (
        Object.prototype.hasOwnProperty.call(
          headerIndex,
          header
        )
      ) {
        newRow.push(
          existingData[row][headerIndex[header]]
        );
      } else {
        newRow.push('');
      }
    });

    newData.push(newRow);
  }

  /*****************************************************
   * CLEAR EXISTING SHEET
   *****************************************************/

  if (sheet.getMaxColumns() < desiredHeaders.length) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      desiredHeaders.length -
        sheet.getMaxColumns()
    );
  }

  sheet.clearContents();

  /*****************************************************
   * WRITE NEW STRUCTURE
   *****************************************************/

  sheet
    .getRange(
      1,
      1,
      newData.length,
      desiredHeaders.length
    )
    .setValues(newData);

  /*****************************************************
   * REMOVE EXTRA COLUMNS
   *****************************************************/

  const extraColumns =
    sheet.getMaxColumns() -
    desiredHeaders.length;

  if (extraColumns > 0) {
    sheet.deleteColumns(
      desiredHeaders.length + 1,
      extraColumns
    );
  }

  /*****************************************************
   * FORMAT
   *****************************************************/

  formatHeader_(sheet);
}

/*******************************************************
 * SAVE LABEL APPLICATION
 *******************************************************/

function saveApplication(data) {
  validateServer_(data);

  const ss = getSpreadsheet_();

  const labelSheet = ensureSheet_(
    ss,
    SHEETS.LABELS,
    LABEL_HEADERS
  );

  const profileSheet = ensureSheet_(
    ss,
    SHEETS.PROFILES,
    PROFILE_HEADERS
  );

  const timestamp = new Date();

  /*****************************************************
   * RECORD ID
   *****************************************************/

  const recordId =
    'TTB-' +
    Utilities.getUuid()
      .replace(/-/g, '')
      .substring(0, 10)
      .toUpperCase();

  /*****************************************************
   * WARNING VALIDATION
   *****************************************************/

  const normalizedWarning =
    normalize_(data.governmentWarning);

  const normalizedRequiredWarning =
    normalize_(WARNING_TEXT);

  const warningMatch =
    normalizedWarning === normalizedRequiredWarning;

  /*****************************************************
   * REVIEW STATUS
   *****************************************************/

  const reviewStatus =
    warningMatch
      ? 'READY FOR HUMAN REVIEW'
      : 'POTENTIAL ISSUE';

  /*****************************************************
   * CLASS/TYPE PROFILE
   *****************************************************/

  const profile = upsertProfile_(
    profileSheet,
    data.beverageCategory,
    data.classType,
    timestamp,
    reviewStatus
  );

  /*****************************************************
   * SAVE APPLICATION
   *****************************************************/

  labelSheet.appendRow([
    recordId,
    timestamp,
    data.beverageCategory || '',
    data.brandName || '',
    data.classType || '',
    data.alcoholContent || '',
    data.alcoholRequired ? 'Yes' : 'No',
    data.netContents || '',
    data.bottlerName || '',
    data.bottlerAddress || '',
    data.producerName || '',
    data.producerAddress || '',
    data.sameAddress ? 'Yes' : 'No',
    data.imported ? 'Yes' : 'No',
    data.countryOfOrigin || '',
    data.governmentWarning || '',
    warningMatch ? 'YES' : 'NO',
    profile.profileId,
    reviewStatus
  ]);

  return {
    success: true,

    recordId: recordId,

    profileId: profile.profileId,

    profileLabel: profile.profileLabel,

    warningProvided: Boolean(
      String(
        data.governmentWarning || ''
      ).trim()
    ),

    warningMatch: warningMatch,

    reviewStatus: reviewStatus,

    timestamp: timestamp.toISOString()
  };
}

/*******************************************************
 * SEARCH CLASS/TYPE PROFILES
 *******************************************************/

function searchProfiles(query) {
  const ss = getSpreadsheet_();

  const sheet = ensureSheet_(
    ss,
    SHEETS.PROFILES,
    PROFILE_HEADERS
  );

  const values =
    sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return [];
  }

  const searchTerm =
    normalize_(query || '');

  const results = [];

  for (
    let i = 1;
    i < values.length;
    i++
  ) {
    const row = values[i];

    const profileId = row[0];
    const classType = row[1];
    const beverageCategory = row[2];
    const submissionCount = row[3];
    const firstSeen = row[4];
    const lastSeen = row[5];
    const lastReviewStatus = row[6];

    const searchableText =
      normalize_([
        profileId,
        classType,
        beverageCategory,
        lastReviewStatus
      ].join(' '));

    if (
      searchTerm === '' ||
      searchableText.indexOf(searchTerm) !== -1
    ) {
      results.push({
        profileId: profileId,
        classType: classType,
        beverageCategory: beverageCategory,
        submissionCount: submissionCount,
        firstSeen: formatDate_(firstSeen),
        lastSeen: formatDate_(lastSeen),
        lastReviewStatus: lastReviewStatus
      });
    }
  }

  /*****************************************************
   * MOST RECENT FIRST
   *****************************************************/

  results.sort(function(a, b) {
    const dateA =
      new Date(
        a.lastSeen || 0
      ).getTime();

    const dateB =
      new Date(
        b.lastSeen || 0
      ).getTime();

    return dateB - dateA;
  });

  return results.slice(0, 100);
}

/*******************************************************
 * GET ALL PROFILES
 *******************************************************/

function getProfiles() {
  return searchProfiles('');
}

/*******************************************************
 * GET SINGLE PROFILE
 *******************************************************/

function getProfile(profileId) {
  const profiles = getProfiles();

  for (
    let i = 0;
    i < profiles.length;
    i++
  ) {
    if (
      profiles[i].profileId === profileId
    ) {
      return profiles[i];
    }
  }

  return null;
}

/*******************************************************
 * DASHBOARD STATISTICS
 *******************************************************/

function getDashboardStats() {
  const ss = getSpreadsheet_();

  const labelSheet = ensureSheet_(
    ss,
    SHEETS.LABELS,
    LABEL_HEADERS
  );

  const profileSheet = ensureSheet_(
    ss,
    SHEETS.PROFILES,
    PROFILE_HEADERS
  );

  const labelData =
    labelSheet
      .getDataRange()
      .getValues();

  const profileData =
    profileSheet
      .getDataRange()
      .getValues();

  let ready = 0;
  let issues = 0;

  /*****************************************************
   * REVIEW STATUS
   *
   * Review Status is column S = index 18
   *****************************************************/

  for (
    let i = 1;
    i < labelData.length;
    i++
  ) {
    const status =
      String(
        labelData[i][18] || ''
      );

    if (
      status ===
      'READY FOR HUMAN REVIEW'
    ) {
      ready++;
    }

    if (
      status ===
      'POTENTIAL ISSUE'
    ) {
      issues++;
    }
  }

  return {
    totalApplications:
      Math.max(
        0,
        labelData.length - 1
      ),

    totalProfiles:
      Math.max(
        0,
        profileData.length - 1
      ),

    readyForReview: ready,

    potentialIssues: issues
  };
}

/*******************************************************
 * RETURN REQUIRED WARNING TEXT
 *******************************************************/

function getWarningText() {
  return WARNING_TEXT;
}

/*******************************************************
 * SERVER-SIDE VALIDATION
 *******************************************************/

function validateServer_(data) {
  if (!data) {
    throw new Error(
      'No application data was received.'
    );
  }

  /*****************************************************
   * BEVERAGE CATEGORY
   *****************************************************/

  if (
    !String(
      data.beverageCategory || ''
    ).trim()
  ) {
    throw new Error(
      'Beverage Category is required.'
    );
  }

  /*****************************************************
   * BRAND NAME
   *****************************************************/

  if (
    !String(
      data.brandName || ''
    ).trim()
  ) {
    throw new Error(
      'Brand Name is required.'
    );
  }

  /*****************************************************
   * CLASS / TYPE
   *****************************************************/

  if (
    !String(
      data.classType || ''
    ).trim()
  ) {
    throw new Error(
      'Class/Type Designation is required.'
    );
  }

  /*****************************************************
   * NET CONTENTS
   *****************************************************/

  if (
    !String(
      data.netContents || ''
    ).trim()
  ) {
    throw new Error(
      'Net Contents is required.'
    );
  }

  /*****************************************************
   * BOTTLER
   *****************************************************/

  if (
    !String(
      data.bottlerName || ''
    ).trim()
  ) {
    throw new Error(
      'Bottler Name is required.'
    );
  }

  if (
    !String(
      data.bottlerAddress || ''
    ).trim()
  ) {
    throw new Error(
      'Bottler Address is required.'
    );
  }

  /*****************************************************
   * PRODUCER
   *****************************************************/

  if (
    !String(
      data.producerName || ''
    ).trim()
  ) {
    throw new Error(
      'Producer Name is required.'
    );
  }

  /*****************************************************
   * PRODUCER ADDRESS
   *
   * Only required when addresses differ.
   *****************************************************/

  if (!data.sameAddress) {
    if (
      !String(
        data.producerAddress || ''
      ).trim()
    ) {
      throw new Error(
        'Producer Address is required when the addresses are different.'
      );
    }
  }

  /*****************************************************
   * IMPORT INFORMATION
   *****************************************************/

  if (data.imported) {
    if (
      !String(
        data.countryOfOrigin || ''
      ).trim()
    ) {
      throw new Error(
        'Country of Origin is required for imported products.'
      );
    }
  }

  /*****************************************************
   * GOVERNMENT HEALTH WARNING
   *****************************************************/

  if (
    !String(
      data.governmentWarning || ''
    ).trim()
  ) {
    throw new Error(
      'Government Health Warning Statement Provided is required.'
    );
  }

  /*****************************************************
   * ALCOHOL CONTENT
   *
   * Beer and wine do not require alcohol content
   * in this prototype.
   *****************************************************/

  if (data.alcoholRequired) {
    if (
      !String(
        data.alcoholContent || ''
      ).trim()
    ) {
      throw new Error(
        'Alcohol Content is required for this beverage category.'
      );
    }
  }
}

/*******************************************************
 * CREATE OR UPDATE CLASS/TYPE PROFILE
 *******************************************************/

function upsertProfile_(
  sheet,
  beverageCategory,
  classType,
  timestamp,
  reviewStatus
) {
  const category =
    String(
      beverageCategory || ''
    ).trim();

  const type =
    String(
      classType || ''
    ).trim();

  const normalizedCategory =
    normalize_(category);

  const normalizedType =
    normalize_(type);

  const profileKey =
    normalizedCategory +
    '|' +
    normalizedType;

  /*****************************************************
   * CREATE PROFILE ID
   *****************************************************/

  const profileId =
    'PROFILE-' +
    (
      normalizedCategory +
      '-' +
      normalizedType
    )
      .replace(/[^A-Z0-9]+/g, '-')
      .replace(/^-|-$/g, '')
      .substring(0, 65);

  /*****************************************************
   * READ EXISTING PROFILES
   *****************************************************/

  const data =
    sheet
      .getDataRange()
      .getValues();

  /*****************************************************
   * SEARCH FOR EXISTING PROFILE
   *****************************************************/

  for (
    let i = 1;
    i < data.length;
    i++
  ) {
    const existingCategory =
      normalize_(data[i][2]);

    const existingType =
      normalize_(data[i][1]);

    const existingKey =
      existingCategory +
      '|' +
      existingType;

    if (
      existingKey === profileKey
    ) {
      const rowNumber = i + 1;

      const currentCount =
        Number(data[i][3]) || 0;

      /***********************************************
       * Increment submission count
       ***********************************************/

      sheet
        .getRange(rowNumber, 4)
        .setValue(
          currentCount + 1
        );

      /***********************************************
       * Update last seen
       ***********************************************/

      sheet
        .getRange(rowNumber, 6)
        .setValue(timestamp);

      /***********************************************
       * Update review status
       ***********************************************/

      sheet
        .getRange(rowNumber, 7)
        .setValue(reviewStatus);

      return {
        profileId: data[i][0],

        profileLabel:
          category +
          ' — ' +
          type
      };
    }
  }

  /*****************************************************
   * CREATE NEW PROFILE
   *****************************************************/

  sheet.appendRow([
    profileId,
    type,
    category,
    1,
    timestamp,
    timestamp,
    reviewStatus
  ]);

  return {
    profileId: profileId,

    profileLabel:
      category +
      ' — ' +
      type
  };
}

/*******************************************************
 * GOOGLE SHEETS CONNECTION
 *******************************************************/

function getSpreadsheet_() {
  return SpreadsheetApp.openById(
    SPREADSHEET_ID
  );
}

/*******************************************************
 * ENSURE SHEET EXISTS
 *******************************************************/

function ensureSheet_(
  spreadsheet,
  sheetName,
  headers
) {
  let sheet =
    spreadsheet.getSheetByName(
      sheetName
    );

  if (!sheet) {
    sheet =
      spreadsheet.insertSheet(
        sheetName
      );
  }

  if (
    sheet.getLastRow() === 0
  ) {
    sheet
      .getRange(
        1,
        1,
        1,
        headers.length
      )
      .setValues([headers]);
  }

  return sheet;
}

/*******************************************************
 * HEADER FORMATTING
 *******************************************************/

function formatHeader_(sheet) {
  const lastColumn =
    sheet.getLastColumn();

  if (lastColumn < 1) {
    return;
  }

  sheet
    .getRange(
      1,
      1,
      1,
      lastColumn
    )
    .setFontWeight('bold');

  sheet.setFrozenRows(1);

  /*****************************************************
   * AUTO RESIZE
   *****************************************************/

  for (
    let i = 1;
    i <= lastColumn;
    i++
  ) {
    sheet.autoResizeColumn(i);
  }
}

/*******************************************************
 * NORMALIZE TEXT
 *
 * Used for comparisons such as:
 *
 * STONE'S THROW
 * Stone's Throw
 *
 * This prevents capitalization differences from
 * automatically creating separate profiles.
 *******************************************************/

function normalize_(value) {
  return String(value || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();
}

/*******************************************************
 * FORMAT DATE
 *******************************************************/

function formatDate_(value) {
  if (!value) {
    return '';
  }

  try {
    return Utilities.formatDate(
      new Date(value),
      Session.getScriptTimeZone(),
      'MM/dd/yyyy HH:mm'
    );
  } catch (error) {
    return String(value);
  }
}

/*******************************************************
 * DATABASE TEST
 *******************************************************/

function testDatabase() {
  const result =
    setupDatabase();

  Logger.log(result);

  return result;
}

/*******************************************************
 * MIGRATION TEST
 *******************************************************/

function testColumnAlignment() {
  const result =
    alignPrototypeColumns();

  Logger.log(result);

  return result;
}
