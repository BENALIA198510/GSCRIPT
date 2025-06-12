/**
 * @file Field Training Management System - Complete Server-Side Implementation
 * @author System Administrator
 * @version 3.1 - Patched for Data Visibility and Optimized Logic
 */
// ======================= CONSTANTS =======================
/**
 * Holds the names of the Google Sheets used by the application.
 * @const {object}
 */
const SHEET_NAMES = {
  LOGIN: 'Login',
  DATA: 'Data',
  OPTIONS: 'Options',
};
/**
 * Maps field names to their respective column indexes (0-based) in Google Sheets.
 * This prevents issues if column order changes.
 * @const {object}
 */
const SHEET_COLUMNS = {
  LOGIN: {
    EMAIL: 0,
    PASSWORD: 1,
    USER_TYPE: 2,
    OTP: 3,
  },
  DATA: {
    SPECIALITE: 0,
    GROUPE: 1,
    NOM_PRENOM: 2,
    CIN: 3,
    DATE_STAGE: 4,
    NOMBRE_HEURES: 5,
    COMMUNE: 6,
    ETABLISSEMENT: 7,
    NOM_ENCADRANT: 8,
    PPR_ENCADRANT: 9,
    RECORD_OWNER: 10,
  },
};
/**
 * General application configuration settings.
 * @const {object}
 */
const CONFIG = {
  DEBUG: false, // Set to true for detailed console logging
  CACHE_DURATION: 300, // Cache duration in seconds (5 minutes)
  OTP_LENGTH: 6, // Length of the One-Time Password
};

// ======================= HELPER UTILITIES =======================
/**
 * Returns the shared script cache instance.
 * @returns {GoogleAppsScript.Cache.Cache}
 */
function getCache() {
  return CacheService.getScriptCache();
}

/**
 * Logs an error and returns a standard error response.
 * @param {string} context - Description of where the error occurred.
 * @param {Error} error - The caught error object.
 * @returns {{success:boolean, message:string}}
 */
function handleServerError(context, error) {
  logError(context, error);
  return {
    success: false,
    message: 'خطأ في الخادم: ' + error.message,
  };
}

/**
 * Retrieves the row index and data for a user based on email.
 * @param {string} email - User email.
 * @returns {{row:number, data:Array}|null}
 */
function getUserRow(email) {
  const sheet = getSheet(SHEET_NAMES.LOGIN);
  const data = sheet.getDataRange().getValues();
  const normalizedEmail = email.trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (data[i][SHEET_COLUMNS.LOGIN.EMAIL] === normalizedEmail) {
      return { row: i + 1, data: data[i] };
    }
  }
  return null;
}

/**
 * Paginates an array of records.
 * @param {Array} list - Full list of records.
 * @param {number} page - Current page number (1-based).
 * @param {number} perPage - Records per page.
 * @returns {Array} Slice of the list for the requested page.
 */
function paginate(list, page, perPage) {
  const start = (page - 1) * perPage;
  return list.slice(start, start + perPage);
}
// ======================= WEB APP ENTRY POINTS =======================
/**
 * Main entry point for the web app when a user accesses the URL.
 * @param {object} e - The event parameter from the HTTP GET request.
 * @returns {HtmlService.HtmlOutput} The HTML service output to be rendered.
 */
function doGet(e) {
  try {
    logDebug('doGet called with parameters:', e.parameter);
    // Serve the main HTML file from a template
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle('نظام إدارة التدريب الميداني');
  } catch (error) {
    logError('Critical doGet error:', error);
    return HtmlService.createHtmlOutput('An error occurred while loading the application.');
  }
}
/**
 * Includes HTML partials into the main HTML file.
 * This allows for modular HTML structure (e.g., separating CSS, JS, or HTML sections).
 * @param {string} filename - The name of the HTML file to include.
 * @returns {string} The content of the requested HTML file.
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    logError(`Include file error for "${filename}":`, error);
    // Return empty string to prevent breaking the main page
    return '';
  }
}
// ======================= SECURITY & AUTHENTICATION =======================
/**
 * Hashes a plain-text password using the SHA-256 algorithm.
 * @param {string} password - The plain text password.
 * @returns {string} The SHA-256 hashed password as a hex string.
 */
function hashPassword(password) {
  try {
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
    // Convert byte array to a hex string
    return digest.map(byte => {
      const unsignedByte = byte < 0 ? byte + 256 : byte;
      return ('0' + unsignedByte.toString(16)).slice(-2);
    }).join('');
  } catch (error) {
    logError('Password hashing error:', error);
    throw new Error('Error occurred during password encryption.');
  }
}
/**
 * Authenticates a user based on email and password.
 * @param {string} email - The user's email address.
 * @param {string} password - The user's plain-text password.
 * @returns {object} An object containing the authentication result.
 */
function userLogin(email, password) {
  try {
    logDebug('Login attempt for:', email);
    if (!email || !password) {
      return {
        success: false,
        message: 'البريد الإلكتروني وكلمة المرور مطلوبان'
      };
    }
    const user = getUserRow(email);
    if (!user) {
      return { success: false, message: 'البريد الإلكتروني غير موجود' };
    }

    const hashedInput = hashPassword(password);
    if (user.data[SHEET_COLUMNS.LOGIN.PASSWORD] === hashedInput) {
      logDebug('Successful login for:', email);
      return {
        success: true,
        userType: user.data[SHEET_COLUMNS.LOGIN.USER_TYPE] || 'User',
      };
    }

    logDebug('Invalid password for:', email);
    return {
      success: false,
      message: 'كلمة المرور غير صحيحة'
    };
  } catch (error) {
    return handleServerError('Login error', error);
  }
}
/**
 * Registers a new user.
 * @param {string} email - The new user's email.
 * @param {string} password - The new user's password.
 * @returns {object} The registration result.
 */
function registerUser(email, password) {
  try {
    logDebug('Registration attempt for:', email);
    if (!email || !password) {
      return {
        success: false,
        message: 'البريد الإلكتروني وكلمة المرور مطلوبان'
      };
    }
    if (password.length < 6) {
      return {
        success: false,
        message: 'كلمة المرور يجب أن تكون 6 أحرف على الأقل'
      };
    }
    const sheet = getSheet(SHEET_NAMES.LOGIN);
    const normalizedEmail = email.trim().toLowerCase();
    if (getUserRow(normalizedEmail)) {
      return {
        success: false,
        message: 'البريد الإلكتروني مستخدم بالفعل'
      };
    }
    // Add the new user to the Login sheet
    const hashedPassword = hashPassword(password);
    sheet.appendRow([normalizedEmail, hashedPassword, 'User', '']); // Default to 'User'
    logDebug('User registered successfully:', email);
    return {
      success: true,
      message: 'تم التسجيل بنجاح'
    };
  } catch (error) {
    return handleServerError('Registration error', error);
  }
}
/**
 * Generates and emails a password reset OTP.
 * @param {string} email - The user's email.
 * @returns {object} The result of the request.
 */
function forgotPasswordRequest(email) {
  try {
    logDebug('Password reset request for:', email);
    if (!email) {
      return {
        success: false,
        message: 'البريد الإلكتروني مطلوب'
      };
    }
    const user = getUserRow(email);
    if (!user) {
      return { success: false, message: 'البريد الإلكتروني غير موجود' };
    }

    const otp = generateOTP();
    const sheet = getSheet(SHEET_NAMES.LOGIN);
    sheet.getRange(user.row, SHEET_COLUMNS.LOGIN.OTP + 1).setValue(otp);

    MailApp.sendEmail({
      to: email,
      subject: 'رمز إعادة تعيين كلمة المرور - نظام التدريب الميداني',
      body: `مرحباً،\n\nرمز التأكيد الخاص بك هو: ${otp}\n\nهذا الرمز صالح لمرة واحدة فقط ولمدة محدودة.\n\nإذا لم تطلب إعادة تعيين كلمة المرور، يرجى تجاهل هذه الرسالة.\n\nشكراً،\nفريق نظام التدريب الميداني`
    });
    logDebug('OTP sent successfully to:', email);
    return { success: true, message: 'تم إرسال رمز التأكيد إلى بريدك الإلكتروني' };
  } catch (error) {
    return handleServerError('Password reset request error', error);
  }
}
/**
 * Verifies the OTP and resets the user's password.
 * @param {string} email - The user's email.
 * @param {string} otp - The OTP code from the email.
 * @param {string} newPassword - The new password.
 * @returns {object} The result of the password reset.
 */
function forgotPasswordVerify(email, otp, newPassword) {
  try {
    logDebug('Password reset verification for:', email);
    if (!email || !otp || !newPassword) {
      return {
        success: false,
        message: 'جميع الحقول مطلوبة'
      };
    }
    if (newPassword.length < 6) {
      return {
        success: false,
        message: 'كلمة المرور يجب أن تكون 6 أحرف على الأقل'
      };
    }
    const user = getUserRow(email);
    if (!user) {
      return { success: false, message: 'رمز التأكيد غير صحيح أو منتهي الصلاحية' };
    }
    const storedOtp = String(user.data[SHEET_COLUMNS.LOGIN.OTP] || '').trim();
    if (storedOtp !== otp.trim()) {
      return { success: false, message: 'رمز التأكيد غير صحيح أو منتهي الصلاحية' };
    }

    const hashedPassword = hashPassword(newPassword);
    const sheet = getSheet(SHEET_NAMES.LOGIN);
    sheet.getRange(user.row, SHEET_COLUMNS.LOGIN.PASSWORD + 1).setValue(hashedPassword);
    sheet.getRange(user.row, SHEET_COLUMNS.LOGIN.OTP + 1).setValue('');
    logDebug('Password reset successful for:', email);
    return { success: true, message: 'تم تغيير كلمة المرور بنجاح' };
  } catch (error) {
    return handleServerError('Password reset verification error', error);
  }
}
/**
 * Checks if a user has 'Admin' privileges.
 * @param {string} email - The email of the user to check.
 * @returns {boolean} True if the user is an admin, false otherwise.
 */
function isAdmin(email) {
  try {
    const user = getUserRow(email);
    return user ? user.data[SHEET_COLUMNS.LOGIN.USER_TYPE] === 'Admin' : false;
  } catch (error) {
    logError('Admin check error:', error);
    return false;
  }
}
// ======================= DATA MANAGEMENT & FILTERING =======================
/**
 * Gets dropdown options for filters, using caching for performance.
 * The data structure is optimized for cascading dropdowns on the frontend.
 * @returns {object} An object containing structured data for dropdowns.
 */
function getDropdownOptions() {
  const cache = getCache();
  const cacheKey = 'dropdownOptions';
  const cached = cache.get(cacheKey);
  if (cached && !CONFIG.DEBUG) {
    logDebug('Returning cached dropdown options');
    return JSON.parse(cached);
  }
  try {
    logDebug('Loading fresh dropdown options from sheet.');
    const sheet = getSheet(SHEET_NAMES.OPTIONS);
    // Get all values, excluding the header row
    const data = sheet.getLastRow() > 1 ?
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues() : [];
    // Use Sets for uniqueness to improve performance over Array.includes() in a loop
    const filterGroup1 = {}; // { specialite: { groupe: Set<nomComplet> } }
    const filterGroup2 = {}; // { commune: { etablissement: Set<nomEncadrant> } }
    data.forEach(([specialite, groupe, nomComplet, commune, etablissement, nomEncadrant]) => {
      // Build student-related filters
      if (specialite && groupe && nomComplet) {
        if (!filterGroup1[specialite]) filterGroup1[specialite] = {};
        if (!filterGroup1[specialite][groupe]) filterGroup1[specialite][groupe] = new Set();
        filterGroup1[specialite][groupe].add(nomComplet);
      }
      // Build establishment-related filters
      if (commune && etablissement && nomEncadrant) {
        if (!filterGroup2[commune]) filterGroup2[commune] = {};
        if (!filterGroup2[commune][etablissement]) filterGroup2[commune][etablissement] = new Set();
        filterGroup2[commune][etablissement].add(nomEncadrant);
      }
    });
    // Convert Sets to Arrays for JSON serialization
    const result = {
      filterGroup1: {},
      filterGroup2: {}
    };
    for (const spec in filterGroup1) {
      result.filterGroup1[spec] = {};
      for (const grp in filterGroup1[spec]) {
        result.filterGroup1[spec][grp] = Array.from(filterGroup1[spec][grp]).sort();
      }
    }
    for (const com in filterGroup2) {
      result.filterGroup2[com] = {};
      for (const etab in filterGroup2[com]) {
        result.filterGroup2[com][etab] = Array.from(filterGroup2[com][etab]).sort();
      }
    }
    cache.put(cacheKey, JSON.stringify(result), CONFIG.CACHE_DURATION);
    logDebug('Dropdown options loaded and cached.');
    return result;
  } catch (error) {
    logError('Error loading dropdown options:', error);
    return {
      filterGroup1: {},
      filterGroup2: {}
    };
  }
}
/**
 * Retrieves and filters training data from the 'Data' sheet.
 * @param {string} email - The current user's email.
 * @param {string} userType - The user's role ('Admin' or 'User').
 * @param {...string} filters - Various filter criteria for the data.
 * @returns {Array<object>} An array of filtered record objects.
 */
// استبدل دالة getData القديمة بهذه النسخة المحسّنة في ملف CODE GS.gs

function getData(email, userType, filterSpecialite, filterGroupe, filterNom, filterCommune, filterEtab, filterEnc, dateFrom, dateTo) {
  try {
    logDebug('getData called for user:', email, 'with filters');
    
    const sheet = getSheet(SHEET_NAMES.DATA);
    const data = sheet.getDataRange().getValues();
    
    const records = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const [specialite, groupe, nomPrenom, cin, dateStage, nombreHeures, commune, etablissement, nomEncadrant, pprEncadrant, recordOwner] = row;
      
      // Skip empty rows
      if (!specialite && !groupe && !nomPrenom) continue;
      
      // Filter by user ownership for non-admin users
      if (userType === 'User' && recordOwner !== email) continue;
      
      // Apply text filters
      if (filterSpecialite && specialite !== filterSpecialite) continue;
      if (filterGroupe && groupe !== filterGroupe) continue;
      if (filterNom && nomPrenom !== filterNom) continue;
      if (filterCommune && commune !== filterCommune) continue;
      if (filterEtab && etablissement !== filterEtab) continue;
      if (filterEnc && nomEncadrant !== filterEnc) continue;
      
      // Apply date filters
      if (dateFrom || dateTo) {
        const recordDate = new Date(dateStage);
        if (isNaN(recordDate.getTime())) continue;
        
        if (dateFrom && recordDate < new Date(dateFrom)) continue;
        if (dateTo && recordDate > new Date(dateTo)) continue;
      }
      
      records.push({
        specialite: specialite || '',
        groupe: groupe || '',
        nomPrenom: nomPrenom || '',
        cin: cin || '',
        dateStage: formatDate(dateStage),
        nombreHeures: nombreHeures || 0,
        commune: commune || '',
        etablissement: etablissement || '',
        nomEncadrant: nomEncadrant || '',
        pprEncadrant: pprEncadrant || '',
        recordOwner: recordOwner || '',
        rowIndex: i + 1
      });
    }
    
    logDebug('Returned', records.length, 'records for user:', email);
    return records;
  } catch (error) {
    logError('Error loading data:', error);
    return [];
  }
}

/**
 * Calculates summary statistics based on user's viewable data.
 * Uses caching to improve performance for repeated calls.
 * @param {string} email - The current user's email.
 * @param {string} userType - The user's role ('Admin' or 'User').
 * @returns {object} An object with summary statistics.
 */
function getSummaryStats(email, userType) {
  const cache = getCache();
  // Admins see all stats, so their cache key is shared.
  // Regular users now also see all data, so everyone can share the admin cache key.
  const cacheKey = 'summary_stats_Admin'; 
  const cached = cache.get(cacheKey);
  if (cached && !CONFIG.DEBUG) {
    logDebug('Returning cached summary stats for:', cacheKey);
    return JSON.parse(cached);
  }
  try {
    // Pass 'Admin' as userType to ensure the getData call fetches all records for stats calculation.
    const records = getData(email, 'Admin', '', '', '', '', '', '', '', '');
    const specialties = new Set();
    const institutions = new Set();
    let totalHours = 0;
    records.forEach(record => {
      if (record.specialite) specialties.add(record.specialite);
      if (record.etablissement) institutions.add(record.etablissement);
      totalHours += Number(record.nombreHeures) || 0;
    });
    const result = {
      totalSpecialties: specialties.size,
      totalStudents: records.length,
      totalHours: totalHours,
      totalInstitutions: institutions.size,
    };
    // Cache the result
    cache.put(cacheKey, JSON.stringify(result), CONFIG.CACHE_DURATION);
    logDebug('Summary stats calculated and cached for:', cacheKey);
    return result;
  } catch (error) {
    logError('Error getting summary stats:', error);
    return {
      totalSpecialties: 0,
      totalStudents: 0,
      totalHours: 0,
      totalInstitutions: 0
    };
  }
}
// ======================= CRUD OPERATIONS =======================
/**
 * Invalidates caches that are affected by data changes.
 * @param {string} email - The email of the user performing the action.
 */
function invalidateCaches(email) {
  const cache = getCache();
  logDebug('Invalidating caches...');
  cache.remove('dropdownOptions'); // Dropdowns might change
  cache.remove('summary_stats_Admin'); // Stats will change for everyone
}
/**
 * Creates a new training record. Admin-only.
 * @param {object} dataObj - The record data from the client.
 * @param {string} email - The email of the admin creating the record.
 * @returns {object} The result of the creation operation.
 */
function createRecord(dataObj, email) {
  try {
    if (!isAdmin(email)) {
      return {
        success: false,
        message: 'ليس لديك صلاحية لإضافة السجلات'
      };
    }
    const validation = validateRecordData(dataObj);
    if (!validation.valid) {
      return {
        success: false,
        message: validation.message
      };
    }
    const sheet = getSheet(SHEET_NAMES.DATA);
    const cinColumnValues = sheet.getLastRow() > 1 ? sheet.getRange(2, SHEET_COLUMNS.DATA.CIN + 1, sheet.getLastRow() - 1, 1).getValues() : [];
    const isDuplicateCIN = cinColumnValues.some(row => row[0].toString().trim() === dataObj.cin.trim());
    if (isDuplicateCIN) {
      return {
        success: false,
        message: 'رقم البطاقة الوطنية موجود مسبقاً'
      };
    }
    sheet.appendRow([
      dataObj.specialite, dataObj.groupe, dataObj.nomPrenom, dataObj.cin.trim(),
      new Date(dataObj.dateStage), Number(dataObj.nombreHeures), dataObj.commune,
      dataObj.etablissement, dataObj.nomEncadrant, dataObj.pprEncadrant.trim(), email, // record owner is the admin who creates it
    ]);
    invalidateCaches(email);
    logDebug('Record created successfully by:', email);
    return {
      success: true,
      message: 'تم إضافة السجل بنجاح'
    };
  } catch (error) {
    return handleServerError('Create record error', error);
  }
}
/**
 * Updates an existing training record. Admin-only.
 * @param {object} dataObj - The updated record data, including rowIndex.
 * @param {string} email - The email of the admin updating the record.
 * @returns {object} The result of the update operation.
 */
function updateRecord(dataObj, email) {
  try {
    if (!isAdmin(email)) {
      return {
        success: false,
        message: 'ليس لديك صلاحية لتعديل السجلات'
      };
    }
    if (!dataObj.rowIndex) {
      return {
        success: false,
        message: 'معرف السجل مطلوب'
      };
    }
    const validation = validateRecordData(dataObj);
    if (!validation.valid) {
      return {
        success: false,
        message: validation.message
      };
    }
    const sheet = getSheet(SHEET_NAMES.DATA);
    const rowIndex = Number(dataObj.rowIndex);
    if (rowIndex < 2 || rowIndex > sheet.getLastRow()) {
      return {
        success: false,
        message: 'السجل غير موجود'
      };
    }
    // Check for duplicate CIN, excluding the current record being edited
    const existingData = sheet.getDataRange().getValues();
    for (let i = 1; i < existingData.length; i++) {
      if ((i + 1 !== rowIndex) && (existingData[i][SHEET_COLUMNS.DATA.CIN].toString().trim() === dataObj.cin.trim())) {
        return {
          success: false,
          message: 'رقم البطاقة الوطنية موجود مسبقاً'
        };
      }
    }
    const rangeToUpdate = sheet.getRange(rowIndex, 1, 1, 11);
    rangeToUpdate.setValues([[
      dataObj.specialite, dataObj.groupe, dataObj.nomPrenom, dataObj.cin.trim(),
      new Date(dataObj.dateStage), Number(dataObj.nombreHeures), dataObj.commune,
      dataObj.etablissement, dataObj.nomEncadrant, dataObj.pprEncadrant.trim(), email,
    ]]);
    invalidateCaches(email);
    logDebug('Record updated successfully by:', email);
    return {
      success: true,
      message: 'تم تحديث السجل بنجاح'
    };
  } catch (error) {
    return handleServerError('Update record error', error);
  }
}
/**
 * Deletes a record from the 'Data' sheet. Admin-only.
 * @param {number} rowIndex - The 1-based row index of the record to delete.
 * @param {string} email - The email of the admin deleting the record.
 * @returns {object} The result of the delete operation.
 */
function deleteRecord(rowIndex, email) {
  try {
    logDebug('Deleting record at row:', rowIndex, 'by user:', email);
    if (!isAdmin(email)) {
      return {
        success: false,
        message: 'ليس لديك صلاحية لحذف السجلات'
      };
    }
    if (!rowIndex || Number(rowIndex) < 2) {
      return {
        success: false,
        message: 'معرف السجل غير صالح'
      };
    }
    const sheet = getSheet(SHEET_NAMES.DATA);
    if (Number(rowIndex) > sheet.getLastRow()) {
      return {
        success: false,
        message: 'السجل غير موجود'
      };
    }
    sheet.deleteRow(Number(rowIndex));
    invalidateCaches(email);
    logDebug('Record deleted successfully by:', email);
    return {
      success: true,
      message: 'تم حذف السجل بنجاح'
    };
  } catch (error) {
    return handleServerError('Delete record error', error);
  }
}
// ======================= EXPORT TO PDF =======================
/**
 * Exports a given set of records to a PDF file and returns a temporary download URL.
 * @param {Array<object>} records - The array of records to export.
 * @returns {string|null} The public URL of the generated PDF, or null on failure.
 */
function exportToPdf(records) {
  try {
    logDebug('Exporting', records.length, 'records to PDF');
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
    const docName = `تصدير_البيانات_${timestamp}`;
    // Create a temporary Google Doc to build the PDF content
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();
    // Set document to Right-to-Left for Arabic
    body.setAttributes({
      [DocumentApp.Attribute.DIRECTION]: DocumentApp.Direction.RIGHT_TO_LEFT
    });
    // Add title and subtitle
    body.appendParagraph('نظام إدارة التدريب الميداني').setHeading(DocumentApp.ParagraphHeading.TITLE).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph(`تصدير البيانات - ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm')}`).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setAttributes({
      [DocumentApp.Attribute.FONT_SIZE]: 12,
      [DocumentApp.Attribute.ITALIC]: true
    });
    body.appendParagraph('');
    if (records.length === 0) {
      body.appendParagraph('لا توجد بيانات للتصدير').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    } else {
      // Add summary and build the table
      body.appendParagraph(`إجمالي السجلات: ${records.length}`).setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setBold(true);
      body.appendParagraph('');
      const headers = ['التخصص', 'المجموعة', 'الاسم الكامل', 'ر.ب.و', 'تاريخ التدريب', 'عدد الساعات', 'الجماعة', 'المؤسسة', 'اسم المشرف', 'ر. تأجير المشرف'];
      const tableData = [headers, ...records.map(r => [r.specialite, r.groupe, r.nomPrenom, r.cin, r.dateStage, r.nombreHeures.toString(), r.commune, r.etablissement, r.nomEncadrant, r.pprEncadrant])];
      const table = body.appendTable(tableData);
      // Style header row
      const headerRowStyle = {};
      headerRowStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#f0f0f0';
      headerRowStyle[DocumentApp.Attribute.BOLD] = true;
      headerRowStyle[DocumentApp.Attribute.FONT_SIZE] = 10;
      table.getRow(0).setAttributes(headerRowStyle);
      // Style data rows
      const dataRowStyle = {};
      dataRowStyle[DocumentApp.Attribute.FONT_SIZE] = 9;
      for (let i = 1; i < table.getNumRows(); i++) {
        table.getRow(i).setAttributes(dataRowStyle);
      }
    }
    doc.saveAndClose();
    // Convert the doc to a PDF blob
    const docFile = DriveApp.getFileById(doc.getId());
    const blob = docFile.getBlob().setName(`${docName}.pdf`);
    const pdfFile = DriveApp.createFile(blob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    // Clean up the temporary document
    docFile.setTrashed(true);
    logDebug('PDF exported successfully:', pdfFile.getUrl());
    return pdfFile.getUrl();
  } catch (error) {
    logError('PDF export error:', error);
    return null;
  }
}
// ======================= UTILITY FUNCTIONS =======================
/**
 * Gets a sheet by name. If it doesn't exist, it creates and initializes it with headers.
 * @param {string} sheetName - The name of the sheet to get or create.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet object.
 */
function getSheet(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    logDebug('Creating missing sheet:', sheetName);
    sheet = spreadsheet.insertSheet(sheetName);
    let headers = [];
    // Initialize headers based on the sheet type
    switch (sheetName) {
      case SHEET_NAMES.LOGIN:
        headers = [
          ['Email', 'Password', 'UserType', 'OTP']
        ];
        break;
      case SHEET_NAMES.DATA:
        headers = [
          ['Specialite', 'Groupe', 'NomPrenom', 'CIN', 'DateStage', 'NombreHeures', 'Commune', 'Etablissement', 'NomEncadrant', 'PprEncadrant', 'RecordOwner']
        ];
        break;
      case SHEET_NAMES.OPTIONS:
        headers = [
          ['Specialite', 'Groupe', 'NomComplet', 'Commune', 'Etablissement', 'NomEncadrant']
        ];
        break;
    }
    if (headers.length > 0) {
      sheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);
    }
  }
  return sheet;
}
/**
 * Generates a random numeric string to be used as an OTP.
 * @returns {string} The generated OTP code.
 */
function generateOTP() {
  return Math.floor(Math.random() * Math.pow(10, CONFIG.OTP_LENGTH))
    .toString()
    .padStart(CONFIG.OTP_LENGTH, '0');
}
/**
 * Formats a date object or string into 'yyyy-MM-dd' format.
 * @param {Date|string} date - The date to format.
 * @returns {string} The formatted date string, or an empty string if invalid.
 */
function formatDate(date) {
  if (!date) return '';
  try {
    const dateObj = date instanceof Date ? date : new Date(date);
    if (isNaN(dateObj.getTime())) return '';
    return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (error) {
    return '';
  }
}
/**
 * Validates the data for a new or updated record.
 * @param {object} dataObj - The record data object.
 * @returns {{valid: boolean, message?: string}} An object indicating if the data is valid.
 */
function validateRecordData(dataObj) {
  const requiredFields = [{
    field: 'specialite',
    name: 'التخصص'
  }, {
    field: 'groupe',
    name: 'المجموعة'
  }, {
    field: 'nomPrenom',
    name: 'الاسم الكامل'
  }, {
    field: 'cin',
    name: 'رقم البطاقة الوطنية'
  }, {
    field: 'dateStage',
    name: 'تاريخ التدريب'
  }, {
    field: 'nombreHeures',
    name: 'عدد الساعات'
  }, {
    field: 'commune',
    name: 'الجماعة'
  }, {
    field: 'etablissement',
    name: 'المؤسسة'
  }, {
    field: 'nomEncadrant',
    name: 'اسم المشرف'
  }, {
    field: 'pprEncadrant',
    name: 'رقم تأجير المشرف'
  }, ];
  for (const {
      field,
      name
    } of requiredFields) {
    if (!dataObj[field] || String(dataObj[field]).trim() === '') {
      return {
        valid: false,
        message: `${name} مطلوب`
      };
    }
  }
  if (isNaN(Number(dataObj.nombreHeures)) || Number(dataObj.nombreHeures) < 1) {
    return {
      valid: false,
      message: 'عدد الساعات يجب أن يكون رقماً أكبر من صفر'
    };
  }
  if (isNaN(new Date(dataObj.dateStage).getTime())) {
    return {
      valid: false,
      message: 'تاريخ التدريب غير صالح'
    };
  }
  // Basic validation for National ID (allows letters and numbers)
  if (!/^[A-Za-z0-9]{1,15}$/.test(dataObj.cin.trim())) {
    return {
      valid: false,
      message: 'تنسيق رقم البطاقة الوطنية غير صالح'
    };
  }
  return {
    valid: true
  };
}
/**
 * Logs messages for debugging, only if CONFIG.DEBUG is true.
 * @param {...any} args - The arguments to log.
 */
function logDebug(...args) {
  if (CONFIG.DEBUG) {
    console.log('[DEBUG]', new Date().toISOString(), ...args);
  }
}
/**
 * Logs error messages to the console.
 * @param {...any} args - The arguments to log as an error.
 */
function logError(...args) {
  console.error('[ERROR]', new Date().toISOString(), ...args);
}
// ======================= SETUP FUNCTION =======================
/**
 * Initializes the application by creating necessary sheets and a default admin user.
 * This function should be run manually once from the script editor.
 */
function setupApplication() {
  try {
    logDebug('Setting up application...');
    // Create all necessary sheets
    Object.values(SHEET_NAMES).forEach(name => getSheet(name));
    // Add a sample admin user if the Login sheet is empty (has only headers)
    const loginSheet = getSheet(SHEET_NAMES.LOGIN);
    if (loginSheet.getLastRow() === 1) {
      const adminPassword = hashPassword('admin123');
      loginSheet.appendRow(['admin@example.com', adminPassword, 'Admin', '']);
      logDebug('Sample admin user created: admin@example.com / admin123');
    }
    logDebug('Application setup completed successfully.');
    return 'تم إعداد التطبيق بنجاح';
  } catch (error) {
    logError('Application setup error:', error);
    return 'خطأ في إعداد التطبيق: ' + error.message;
  }
}
