/**
 * @OnlyCurrentDoc
 */

const ADDON_NAME = "差し込み";

/**
 * Creates a custom menu in Google Sheets when the spreadsheet opens.
 */
function onOpen() {
    DocumentApp.getUi()
        .createMenu(ADDON_NAME)
        .addItem("Start", showPickerDialog.name)
        .addToUi();
}

/**
 * Displays an HTML-service dialog in Google Sheets that contains client-side
 * JavaScript code for the Google Picker API.
 */
function showPickerDialog() {
    const htmlTemplate = HtmlService.createTemplateFromFile("PickerDialog.html");

    const htmlOutput = htmlTemplate
        .evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setWidth(1099)
        .setHeight(698)
        .addMetaTag("viewport", "width=device-width, initial-scale=1");

    DocumentApp.getUi().showModalDialog(htmlOutput, "Select a file");
}
/**
 * Checks that the file can be accessed.
 */
function getFile(fileId) {
    return Drive.Files.get(fileId, { fields: "*" });
}

function getSpreadsheet(spreadsheetId) {
    const fields = 'spreadsheetId,properties,sheets(properties),spreadsheetUrl';

    return Sheets.Spreadsheets.get(spreadsheetId, {
        fields: fields,
        includeGridData: false,
        excludeTablesInBandedRanges: false,
    });
}

function getSheetsData(fileId, sheetId) {
    const fields = 'sheets(data(rowData(values(formattedValue))))';

    return Sheets.Spreadsheets.getByDataFilter(fileId, {
        fields: fields,
        dataFilters: [{
            gridRange: {
                sheetId: sheetId,
            }
        }],
        // includeGridData: true,
        excludeTablesInBandedRanges: false,
    });
}

function openMergeDialog(context) {
    console.log(context);
    console.log(JSON.stringify(context));

    const htmlTemplate = HtmlService.createTemplateFromFile('MergeDialog.html');
    htmlTemplate.context = JSON.stringify(context);

    const htmlOutput = htmlTemplate.evaluate();

    const ui = DocumentApp.getUi();
    ui.showModalDialog(htmlOutput, ADDON_NAME);

    return context;
}

function openSidebar(context) {
    console.log(context);

    const htmlTemplate = HtmlService.createTemplateFromFile('Sidebar.html');
    htmlTemplate.context = JSON.stringify(context);

    const htmlOutput = htmlTemplate.evaluate();

    const ui = DocumentApp.getUi();
    ui.showSidebar(htmlOutput);

    return context;
}

/**
 * Gets the user's OAuth 2.0 access token so that it can be passed to Picker.
 * This technique keeps Picker from needing to show its own authorization
 * dialog, but is only possible if the OAuth scope that Picker needs is
 * available in Apps Script. In this case, the function includes an unused call
 * to a DriveApp method to ensure that Apps Script requests access to all files
 * in the user's Drive.
 *
 * @return {string} The user's OAuth 2.0 access token.
 */
function getOAuthToken() {
    return ScriptApp.getOAuthToken();
}

function getApiKey() {
    return PropertiesService.getScriptProperties().getProperty("API_KEY");
}

function getAppId() {
    return PropertiesService.getScriptProperties().getProperty("APP_ID");
}
