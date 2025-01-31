function getHubSpotProperties() {
  const token = 'HubSpot_API_Token';
  const objects = ['company', 'contact', 'deal', 'ticket', 'email' , 'events', 'partnerships', 'call', 'feedback_submission', 'marketing_event', 'meeting', 'plans', 'product', 'quote', 'referrals', 'user', 'automation_platform_flow', 'line_item', 'tasks', '2-37766506', '2-39538272'];
  const objectMapping = {
    '2-37766506': 'Chargebee Invoices',
    '2-39538272': 'Demo Meetings'
  };
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Properties");
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
        prop["Object Type"] = objectMapping[objectType] || objectType;
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