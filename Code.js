// ============ CONFIGURATION ============
const LOOKUP_SHEET = "Lookup";
const CATEGORIES = ["個人", "一齊", "交通", "公司/學校", "收入"];


// ============ GET FAVOURITES ============
function getFavourites() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOOKUP_SHEET);
  if (!sheet) return { his: [], hers: [] };
  
  const data = sheet.getDataRange().getValues();
  const his = [];
  const hers = [];
  
  // Look for headers: Person, Title, Price, Category, Route
  // Assuming the sheet has these columns in order
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue; // Skip empty rows
    
    const person = row[0] ? row[0].toString().trim() : '';
    const title = row[1] ? row[1].toString().trim() : '';
    const price = parseFloat(row[2]) || 0;
    const category = row[3] ? row[3].toString().trim() : '交通';
    const route = row[4] ? row[4].toString().trim() : '';
    
    if (!title) continue;
    
    const fav = {
      title: title,
      price: price,
      category: category,
      route: route
    };
    
    if (person === '戰神') {
      his.push(fav);
    } else if (person === '紫璃') {
      hers.push(fav);
    }
  }
  
  return { his, hers };
}

// ============ ADD FAVOURITE (Optional) ============
function addFavourite(person, title, price, category, route) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(LOOKUP_SHEET);
    
    if (!sheet) {
      sheet = ss.insertSheet(LOOKUP_SHEET);
      const headers = ["Person", "Title", "Price", "Category", "Route"];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    
    const lastRow = sheet.getLastRow();
    const newRow = Math.max(lastRow + 1, 2);
    
    const rowData = [
      person === 'his' ? '戰神' : '紫璃',
      title,
      parseFloat(price) || 0,
      category || '交通',
      route || ''
    ];
    
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    
    return { 
      success: true, 
      message: '✅ Favourite added!'
    };
  } catch (error) {
    return { success: false, message: '❌ Error: ' + error.toString() };
  }
}

// ============ DELETE FAVOURITE (Optional) ============
function deleteFavourite(person, title) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LOOKUP_SHEET);
    if (!sheet) return { success: false, message: 'Sheet not found' };
    
    const data = sheet.getDataRange().getValues();
    const personName = person === 'his' ? '戰神' : '紫璃';
    
    for (let i = data.length - 1; i > 0; i--) {
      if (data[i][0] === personName && data[i][1] === title) {
        sheet.deleteRow(i + 1);
        return { success: true, message: '🗑️ Favourite deleted!' };
      }
    }
    
    return { success: false, message: 'Favourite not found' };
  } catch (error) {
    return { success: false, message: '❌ Error: ' + error.toString() };
  }
}

// ============ DO GET ============
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("💰 Our Finance Tracker")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// ============ GET MONTHS ============
function getMonths() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const months = [];
  for (let i = 0; i < sheets.length; i++) {
    const name = sheets[i].getName();
    if (name !== LOOKUP_SHEET && name !== "Summary") {
      months.push(name);
    }
  }
  return months.sort();
}

// ============ GET ROUTES FOR LOOKUP ============
function getRoutes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOOKUP_SHEET);
  if (!sheet) return {};
  
  const data = sheet.getDataRange().getValues();
  const routes = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      routes[data[i][0].toString().trim()] = data[i][1] || data[i][0];
    }
  }
  return routes;
}

// ============ GET ENTRIES ============
function getEntries(month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(month);
  if (!sheet) return { his: [], hers: [] };
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { his: [], hers: [] };
  
  const his = [];
  const hers = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    
    const entry = {
      date: row[0] || "",
      person: row[1] || "",
      category: row[2] || "",
      title: row[3] || "",
      price: parseFloat(row[4]) || 0,
      route: row[5] || "",
      rowIndex: i
    };
    
    if (entry.person === '戰神') {
      his.push(entry);
    } else if (entry.person === '紫璃') {
      hers.push(entry);
    }
  }
  
  return { his, hers };
}

// ============ ADD ENTRY ============
function addEntry(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(data.month);
    
    if (!sheet) {
      sheet = ss.insertSheet(data.month);
      const headers = ["Date", "Person", "Category", "Title", "Price", "Route"];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    
    const lastRow = sheet.getLastRow();
    const newRow = Math.max(lastRow + 1, 2);
    
    let route = "";
    if (data.category === "交通" && data.routeInput) {
      const routes = getRoutes();
      route = routes[data.routeInput] || data.routeInput;
    }
    
    const rowData = [
      data.date,
      data.person === 'his' ? '戰神' : '紫璃',
      data.category,
      data.title,
      parseFloat(data.price) || 0,
      route
    ];
    
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    
    return { 
      success: true, 
      message: `✅ ${data.person === 'his' ? '戰神' : '紫璃'}'s entry added!`,
      route: route
    };
  } catch (error) {
    return { success: false, message: "❌ Error: " + error.toString() };
  }
}

// ============ DELETE ENTRY ============
function deleteEntry(month, rowIndex, person) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(month);
    if (!sheet) return { success: false, message: "Sheet not found" };
    sheet.deleteRow(rowIndex);
    return { success: true, message: "🗑️ Deleted!" };
  } catch (error) {
    return { success: false, message: "❌ Error: " + error.toString() };
  }
}

// ============ GET SUMMARY ============
function getSummary(month) {
  const data = getEntries(month);
  const summary = {
    hisTotal: 0,
    hersTotal: 0,
    togetherTotal: 0,
    hisIncome: 0,
    hersIncome: 0,
    hisByCategory: {},
    hersByCategory: {}
  };
  
  data.his.forEach(entry => {
    const price = parseFloat(entry.price) || 0;
    if (entry.category === "收入") {
      summary.hisIncome += price;
    } else {
      summary.hisTotal += price;
      summary.hisByCategory[entry.category] = (summary.hisByCategory[entry.category] || 0) + price;
      if (entry.category === "一齊") {
        summary.togetherTotal += price;
      }
    }
  });
  
  data.hers.forEach(entry => {
    const price = parseFloat(entry.price) || 0;
    if (entry.category === "收入") {
      summary.hersIncome += price;
    } else {
      summary.hersTotal += price;
      summary.hersByCategory[entry.category] = (summary.hersByCategory[entry.category] || 0) + price;
      if (entry.category === "一齊") {
        summary.togetherTotal += price;
      }
    }
  });
  
  return summary;
}

// ============ GET AVAILABLE YEARS ============
function getYears() {
  const years = {};
  getMonths().forEach(month => {
    const data = getEntries(month);
    [...data.his, ...data.hers].forEach(entry => {
      const year = getEntryYear(entry.date);
      if (year) years[year] = true;
    });
  });
  return Object.keys(years).sort();
}

function getEntryYear(dateValue) {
  if (!dateValue) return null;
  if (dateValue instanceof Date && !isNaN(dateValue.getTime())) {
    return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), "yyyy");
  }

  const dateText = dateValue.toString().trim();
  const yearMatch = dateText.match(/^(\d{4})[-\/]/);
  if (yearMatch) return yearMatch[1];

  const parsedDate = new Date(dateText);
  return isNaN(parsedDate.getTime()) ? null : parsedDate.getFullYear().toString();
}

// ============ GET YEAR SUMMARY ============
function getYearSummary(year) {
  const selectedYear = year.toString();
  const summary = {
    year: selectedYear,
    hisTotal: 0,
    hersTotal: 0,
    togetherTotal: 0,
    hisIncome: 0,
    hersIncome: 0,
    hisByCategory: {},
    hersByCategory: {}
  };

  getMonths().forEach(month => {
    const data = getEntries(month);
    [...data.his, ...data.hers].forEach(entry => {
      if (getEntryYear(entry.date) !== selectedYear) return;

      const price = parseFloat(entry.price) || 0;
      const isHis = entry.person === "戰神";
      const byCategory = isHis ? summary.hisByCategory : summary.hersByCategory;

      if (entry.category === "收入") {
        if (isHis) summary.hisIncome += price;
        else summary.hersIncome += price;
        return;
      }

      if (isHis) summary.hisTotal += price;
      else summary.hersTotal += price;
      byCategory[entry.category] = (byCategory[entry.category] || 0) + price;
      if (entry.category === "一齊") summary.togetherTotal += price;
    });
  });

  return summary;
}