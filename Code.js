/**
 * @OnlyCurrentDoc
 */

/**
 * On open
 * @param {GoogleAppsScript.Events.DocsOnOpen} event
 */
function onOpen(event) {
    createMenu();
}

/**
 * On install
 * @param {GoogleAppsScript.Events.AppsScriptEvent} event
 */
function onInstall(event) {
    createMenu(event);
}

/**
 * Create menu
 * @param {GoogleAppsScript.Events.AppsScriptEvent} event
 */
function createMenu(event) {
    DocumentApp.getUi()
        .createMenu("差し込み文書")
        .addItem("データソースを選択", openPickerDialog.name)
        .addToUi();
}

/**
 * Open picker dialog
 */
function openPickerDialog() {
    const htmlTemplate = HtmlService.createTemplateFromFile("PickerDialog.html");

    const htmlOutput = htmlTemplate.evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setWidth(1099)
        .setHeight(698)
        .addMetaTag("viewport", "width=device-width, initial-scale=1");

    const ui = DocumentApp.getUi();
    ui.showModalDialog(htmlOutput, "データソースを選択");
}

/**
 * Open sidebar
 * @param {Object} context
 */
function openSidebar(context) {
    const htmlTemplate = HtmlService.createTemplateFromFile('Sidebar.html');
    htmlTemplate.context = JSON.stringify(context);

    const htmlOutput = htmlTemplate.evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .addMetaTag("viewport", "width=device-width, initial-scale=1")
        .setTitle('設定');

    const ui = DocumentApp.getUi();
    ui.showSidebar(htmlOutput);
}

/**
 * Open merge dialog
 * @param {Object} context
 */
function openMergeDialog(context) {
    const htmlTemplate = HtmlService.createTemplateFromFile('MergeDialog.html');
    htmlTemplate.context = JSON.stringify(context);

    const htmlOutput = htmlTemplate.evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setWidth(1099)
        .setHeight(698)
        .addMetaTag("viewport", "width=device-width, initial-scale=1");

    const ui = DocumentApp.getUi();
    ui.showModalDialog(htmlOutput, "差し込むデータを選択");
}

/**
 * Create target folder
 * @param {string} templateDocumentId - template document id
 * @param {string} [parentId] - parent folder id (optional)
 * @returns {Object} Target folder
 */
function createTargetFolder(templateDocumentId, parentId) {
    const templateDocument = DriveApp.getFileById(templateDocumentId);
    const parentFolders = [];
    if (parentId) {
        const parentFolder = DriveApp.getFolderById(parentId);
        parentFolders.push(parentFolder);

    } else {
        parentFolders.push(...getParentsFolders_(templateDocument));
    }

    if (parentFolders.length === 0) {
        parentFolders.push(DriveApp.getRootFolder());
    }

    const availableFolder = parentFolders[0];
    if (parentFolders.length > 1) {
        console.log('Multiple parent folders found. Using the first one.');
    }

    let targetFolder;

    try {
        targetFolder = availableFolder.createFolder(`[差し込み文書]${templateDocument.getName()}`);
    } catch (error) {
        targetFolder = DriveApp.getRootFolder()
    }

    return {
        name: targetFolder.getName(),
        url: targetFolder.getUrl(),
        id: targetFolder.getId(),
    };
}

/**
 * Create merge document
 * @param {string} templateDocumentId
 * @param {string} targetFolderId
 * @param {string} name
 * @returns {Object} Merge document
 */
function createMergeDocument(templateDocumentId, targetFolderId, name) {
    const templateFile = DriveApp.getFileById(templateDocumentId);
    const targetFolder = DriveApp.getFolderById(targetFolderId);
    const mergeFile = templateFile.makeCopy(name, targetFolder);
    const url = mergeFile.getUrl();
    const mergeDocument = Docs.Documents.get(mergeFile.getId(), {
        includeTabsContent: false,
    });

    return {
        url: url,
        document: mergeDocument,
    };
}

/**
 * Create and clear merge document
 * @param {string} templateDocumentId
 * @param {string} targetFolderId
 * @param {string} name
 * @returns {Object} Merge document
 */
function createAndClearMergeDocument(templateDocumentId, targetFolderId, name) {
    const templateFile = DriveApp.getFileById(templateDocumentId);
    const targetFolder = DriveApp.getFolderById(targetFolderId);
    const mergeFile = templateFile.makeCopy(name, targetFolder);

    const docs = DocumentApp.openById(mergeFile.getId());
    docs.getBody().clear();
    docs.saveAndClose();

    const url = mergeFile.getUrl();
    const mergeDocument = Docs.Documents.get(mergeFile.getId(), {
        includeTabsContent: false,
    });

    return {
        url: url,
        document: mergeDocument,
    };
}

/**
 * Replace merge document
 * @param {string} documentId
 * @param {string} revisionId
 * @param {Array<GoogleAppsScript.Docs.Schema.Request>} replaceAllTextRequests
 * @returns {GoogleAppsScript.Docs.Schema.BatchUpdateResponse} Response
 */
function replaceMergeDocument(documentId, revisionId, replaceAllTextRequests) {
    const response = Docs.Documents.batchUpdate({
        requests: replaceAllTextRequests,
        writeControl: {
            requiredRevisionId: revisionId,
        }
    }, documentId);

    return response;
}

/**
 * Get template document
 * @returns {GoogleAppsScript.Docs.Schema.Document} Template document
 */
function getTemplateDocument() {
    const activeDocument = DocumentApp.getActiveDocument();
    const templateDocument = Docs.Documents.get(activeDocument.getId(), {
        includeTabsContent: false,
    });

    return templateDocument;
}

/**
 * Replace and append document
 * @param {string} targetDocumentId
 * @param {Object} mergeData
 */
function replaceAndAppendDocument(targetDocumentId, mergeData) {
    const targetDocument = DocumentApp.openById(targetDocumentId);
    const targetDocumentBody = targetDocument.getBody();
    const templateDocumentBodyCopy = DocumentApp.getActiveDocument()
        .getBody()
        .copy();

    for (const fieldCode in mergeData) {
        const replaceText = mergeData[fieldCode];
        templateDocumentBodyCopy.replaceText(fieldCode, replaceText);
    }

    const numChildren = templateDocumentBodyCopy.getNumChildren();
    for (let index = 0; index < numChildren; index++) {
        const childElement = templateDocumentBodyCopy.getChild(index);
        const elementType = childElement.getType();

        switch (elementType) {
            case DocumentApp.ElementType.HORIZONTAL_RULE:
                targetDocumentBody.appendHorizontalRule(childElement.asHorizontalRule().copy());
                break;
            case DocumentApp.ElementType.INLINE_IMAGE:
                targetDocumentBody.appendImage(childElement.asInlineImage().copy());
                break;
            case DocumentApp.ElementType.LIST_ITEM:
                targetDocumentBody.appendListItem(childElement.asListItem().copy());
                break;
            case DocumentApp.ElementType.PAGE_BREAK:
                targetDocumentBody.appendPageBreak(childElement.asPageBreak().copy());
                break;
            case DocumentApp.ElementType.PARAGRAPH:
                targetDocumentBody.appendParagraph(childElement.asParagraph().copy());
                break;
            case DocumentApp.ElementType.TABLE:
                targetDocumentBody.appendTable(childElement.asTable().copy());
                break;
            default:
                console.log('Unknown element type: ' + elementType);
        }
    }

    targetDocument.saveAndClose();
}

/**
 * Get spreadsheet
 * @param {string} spreadsheetId - spreadsheet ID
 * @returns {GoogleAppsScript.Sheets.Schema.Spreadsheet} Spreadsheet
 */
function getSpreadsheet(spreadsheetId) {
    const fields = 'spreadsheetId,properties,sheets(properties),spreadsheetUrl';
    const spreadsheet = Sheets.Spreadsheets.get(spreadsheetId, {
        fields: fields,
        includeGridData: false,
        excludeTablesInBandedRanges: false,
    });
    return spreadsheet;
}

/**
 * Get sheet values
 * @param {string} spreadsheetId - spreadsheet ID
 * @param {number} sheetId - sheet ID
 * @param {number} startRowIndex - start row index
 * @param {number} endRowIndex - end row index
 * @param {number} startColumnIndex - start column index
 * @param {number} endColumnIndex - end column index
 * @returns {Array<Array<string>>} Sheet values
 */
function getSheetValues(spreadsheetId, sheetId, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex) {
    const gridRange = {
        sheetId: sheetId,
        startRowIndex: startRowIndex,
        endRowIndex: endRowIndex,
        startColumnIndex: startColumnIndex,
        endColumnIndex: endColumnIndex,
    };
    const request = {
        dataFilters: [{
            gridRange: gridRange
        }],
        majorDimension: 'ROWS',
        valueRenderOption: 'FORMATTED_VALUE',
        dateTimeRenderOption: 'FORMATTED_STRING',
    };
    const response = Sheets.Spreadsheets.Values.batchGetByDataFilter(request, spreadsheetId);
    return response;
}

/**
 * Get parents folders
 * @param {GoogleAppsScript.Drive.File} file - target file
 * @returns {Array<GoogleAppsScript.Drive.Folder>} Parent folders
 */
function getParentsFolders_(file) {
    const parentFolders = [];
    const folders = file.getParents();
    while (folders.hasNext()) {
        const folder = folders.next();
        parentFolders.push(folder);
    }
    return parentFolders;
}

/**
 * (deprecated) Select available folders 
 * @param {Array<GoogleAppsScript.Drive.Folder>} folders - parent folders
 * @returns {Array<GoogleAppsScript.Drive.Folder>} Available folders
 */
function selectAvailableFolders_(folders) {
    const user = Session.getActiveUser();
    const availableFolders = [];
    for (const folder of folders) {
        const permission = folder.getAccess(user);
        if (permission === DriveApp.Permission.EDIT ||
            permission === DriveApp.Permission.OWNER ||
            permission === DriveApp.Permission.ORGANIZER ||
            permission === DriveApp.Permission.FILE_ORGANIZER) {
            availableFolders.push(folder);
        } else {
            console.log(`No permission to write: ${folder.getName()}`);
            console.log(permission.toString());
        }
    }
    if (availableFolders.length === 0) {
        availableFolders.push(DriveApp.getRootFolder());
    }

    return availableFolders;
}

/**
 * Include HTML file
 * @param {string} filename - HTML file name
 * @returns {string} HTML content
 */
function include_(filename) {
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent();
}

/**
 * Get OAuth token
 * @returns {string} OAuth token
 */
function getOAuthToken() {
    return ScriptApp.getOAuthToken();
}

/**
 * Get API key
 * @returns {string} API key
 */
function getApiKey() {
    return PropertiesService.getScriptProperties()
        .getProperty("API_KEY");
}

/**
 * Get app ID
 * @returns {string} App ID
 */
function getAppId() {
    return PropertiesService.getScriptProperties()
        .getProperty("APP_ID");
}
