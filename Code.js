const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Fetch Rates for the Frontend Dropdown
function getRateMaster() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Rate_Master');
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  return data.map(row => {
    let obj = {};
    headers.forEach((header, i) => obj[header] = row[i]);
    return obj;
  });
}

// Core Calculation Engine (Server-side for Security)
function calculateTotals(qty, rate, labPct, ovhPct, marginPct) {
  const materialCost = qty * rate;
  const labourCost = materialCost * (labPct / 100);
  const overheadCost = (materialCost + labourCost) * (ovhPct / 100);
  const totalCost = materialCost + labourCost + overheadCost;
  const finalPrice = totalCost * (1 + marginPct / 100);

  return {
    materialCost: materialCost.toFixed(2),
    labourCost: labourCost.toFixed(2),
    overheadCost: overheadCost.toFixed(2),
    totalCost: totalCost.toFixed(2),
    finalPrice: finalPrice.toFixed(2)
  };
}

// Save Estimate Data
function saveEstimate(seplId, itemsArray) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Estimate_Items');
  itemsArray.forEach(item => {
    sheet.appendRow([
      seplId, 1, item.code, item.desc, item.qty, item.unit, item.rate, 
      item.matCost, item.labPct, item.labCost, item.ovhPct, item.ovhCost, 
      item.totalCost, item.marginPct, item.finalPrice
    ]);
  });
  return "Success";
}
function getRateByCode(itemCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Rate_Master');
  const data = sheet.getDataRange().getValues();
  
  // Look for the code in the first column
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === itemCode) {
      return {
        description: data[i][2],
        unit: data[i][3],
        rate: data[i][5]
      };
    }
  }
  return null; // If not found
}
function getUserRole() {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email && data[i][3] === true) {
      return {
        email: email,
        role: data[i][1], // Director, Manager, or User
        name: data[i][2]
      };
    }
  }
  return { role: "Guest" }; // If not in the list
}
function checkUserAccess(email, role) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    // Check if Email matches AND Role matches
    if (data[i][0].toLowerCase() === email.toLowerCase() && data[i][1] === role) {
      return {
        success: true,
        name: data[i][2]
      };
    }
  }
  return { success: false };
}