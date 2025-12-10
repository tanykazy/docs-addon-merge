/**
 * @OnlyCurrentDoc
 */

function onOpen() {
    DocumentApp.getUi()
        .createMenu("差し込み")
        .addItem("データソースを選択", openPickerDialog.name)
        .addToUi();
}

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

function insertFieldCode(fieldCode) {
    const document = DocumentApp.getActiveDocument();
    const cursor = document.getCursor();
    cursor.insertText(fieldCode);
}

function createTargetFolder(templateDocumentId) {
    const templateDocument = DriveApp.getFileById(templateDocumentId);
    const parentFolders = getParentsFolders_(templateDocument);
    const availableFolders = selectAvailableFolders_(parentFolders);
    const availableFolder = availableFolders[0];
    if (availableFolders.length > 1) {
        console.log('Multiple available folders found. Using the first one.');
    }
    const targetFolder = DriveApp.createFolder(`[Sashikomi]${templateDocument.getName()}`);
    targetFolder.moveTo(availableFolder);

    return {
        name: targetFolder.getName(),
        url: targetFolder.getUrl(),
        id: targetFolder.getId(),
    };
}

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

function replaceMergeDocument(documentId, revisionId, replaceAllTextRequests) {
    const response = Docs.Documents.batchUpdate({
        requests: replaceAllTextRequests,
        writeControl: {
            requiredRevisionId: revisionId,
        }
    }, documentId);
    return response;
}

function getTemplateDocument() {
    const activeDocument = DocumentApp.getActiveDocument();
    const templateDocument = Docs.Documents.get(activeDocument.getId(), {
        includeTabsContent: false,
    });
    return templateDocument;
}

function getParentsFolders_(file) {
    const parentFolders = [];
    const folders = file.getParents();
    while (folders.hasNext()) {
        const folder = folders.next();
        parentFolders.push(folder);
    }
    if (parentFolders.length === 0) {
        parentFolders.push(DriveApp.getRootFolder());
    }

    return parentFolders;
}

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
        }
    }
    if (availableFolders.length === 0) {
        availableFolders.push(DriveApp.getRootFolder());
    }

    return availableFolders;
}

function getSpreadsheet(spreadsheetId) {
    const fields = 'spreadsheetId,properties,sheets(properties),spreadsheetUrl';
    const response = Sheets.Spreadsheets.get(spreadsheetId, {
        fields: fields,
        includeGridData: false,
        excludeTablesInBandedRanges: false,
    });
    return response;
}

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

function getOAuthToken() {
    return ScriptApp.getOAuthToken();
}

function getApiKey() {
    return PropertiesService.getScriptProperties()
        .getProperty("API_KEY");
}

function getAppId() {
    return PropertiesService.getScriptProperties()
        .getProperty("APP_ID");
}
