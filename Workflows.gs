function getHubSpotFlowsWithFullDetails() {
  const ACCESS_TOKEN = "HubSpot_API_Token";
  const BULK_API_URL = "https://api.hubapi.com/automation/v4/flows";
  const SINGLE_FLOW_API_URL = "https://api.hubapi.com/automation/v4/flows";
  const SHEET_NAME = "Workflows";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME) 
                || SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);
  sheet.clear();
  const bulkWorkflows = [];
  let after = null;
  do {
    const url = `${BULK_API_URL}?limit=100${after ? `&after=${after}` : ""}`;
    const response = UrlFetchApp.fetch(url, {
      method: "GET",
      headers: {
        "Authorization": `Bearer ${ACCESS_TOKEN}`,
        "Content-Type": "application/json"
      }
    });
    const data = JSON.parse(response.getContentText());
    if (data.results && data.results.length > 0) {
      bulkWorkflows.push(...data.results);
      after = data.paging?.next?.after || null;
    } else {
      after = null;
    }
  } while (after);
  if (bulkWorkflows.length === 0) {
    sheet.appendRow(["No workflows found!"]);
    return;
  }
  const allDetails = [];
  const headersSet = new Set();
  bulkWorkflows.forEach(flow => {
    const workflowId = flow.id;
    const detailedFlow = getWorkflowDetails(SINGLE_FLOW_API_URL, workflowId, ACCESS_TOKEN);
    const combinedFlow = { ...flow, ...detailedFlow };
    addKeysToSet(combinedFlow, "", headersSet);
    allDetails.push(combinedFlow);
  });
  const headers = Array.from(headersSet);
  sheet.appendRow(headers);
  allDetails.forEach(details => {
    const row = headers.map(header => {
      const value = getValueByPath(details, header);
      if(value !== undefined)
      {
        value = JSON.stringify(value)
        if (typeof value === "string" && value.startsWith('"') && value.endsWith('"'))
        {
          value = value.slice(1, -1);
        }
      }
      else
      {
        value = "";
      }
      return value;
    });
    sheet.appendRow(row);
  });
  var range = sheet.getDataRange();
  range.setBorder(true, true, true, true, true, true);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
}
function getWorkflowDetails(apiUrl, workflowId, accessToken) {
  const url = `${apiUrl}/${workflowId}`;
  const response = UrlFetchApp.fetch(url, {
    method: "GET",
    headers: {
      "Authorization": `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    }
  });
  return JSON.parse(response.getContentText());
}
function addKeysToSet(obj, prefix, set) {
  for (const key in obj) {
    const fullKey = prefix ? `${prefix}.${key}` : key;
    set.add(fullKey);

    if (typeof obj[key] === "object" && obj[key] !== null && !Array.isArray(obj[key])) {
      addKeysToSet(obj[key], fullKey, set);
    }
  }
}
function getValueByPath(obj, path) {
  return path.split(".").reduce((acc, key) => (acc && acc[key] !== undefined ? acc[key] : undefined), obj);
}