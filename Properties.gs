function getHubSpotProperties() {
  const token = 'HubSpot_API_Token';
  const objects = ['company', 'contact', 'deal', 'ticket', 'events', 'partnerships', 'call', 'feedback_submission', 'marketing_event', 'meeting', 'plans', 'product', 'quote', 'referrals', 'user', 'automation_platform_flow', 'line_item', 'tasks'];
  const SHEET_NAME = "Properties";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME) 
                || SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);
  sheet.clear();
  let allHeaders = new Set();
  let allData = [];
  objects.forEach(function(objectType) {
    const url = `https://api.hubapi.com/crm/v3/properties/${objectType}`;
    const options = {
      'method': 'get',
      'headers': {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    };
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    if (data.results && data.results.length > 0) {
      data.results.forEach(function(prop) {
        Object.keys(prop).forEach(key => allHeaders.add(key));
        prop["Object Type"] = objectType;
        allData.push(prop);
      });
    }
  });
  const headers = ["Object Type", ...Array.from(allHeaders)];
  sheet.appendRow(headers);
  allData.forEach(function(prop) {
    const row = headers.map(header => {
      let cellValue = prop[header] || "";
      if (header === 'options') {
        if (Array.isArray(cellValue) && cellValue.length > 0) {
          return JSON.stringify(cellValue);
        } else {
          return "[]";
        }
      }
      if (Array.isArray(cellValue)) {
        return cellValue.join(", ");
      } else if (typeof cellValue === 'object' && cellValue !== null) {
        return JSON.stringify(cellValue);
      } else {
        return cellValue;
      }
    });
    sheet.appendRow(row);
  });
  var range = sheet.getDataRange();
  range.setBorder(true, true, true, true, true, true);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  sheet.getRange(1, 1, sheet.getLastRow(), 1).createFilter();
}