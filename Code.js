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