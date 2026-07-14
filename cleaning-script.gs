function onFormSubmit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rawSheet = ss.getSheetByName("Form Responses 1");
  var cleanSheet = ss.getSheetByName("Cleaned Data");

  var lastRow = rawSheet.getLastRow();
  var rawData = rawSheet.getRange(lastRow, 1, 1, 10).getValues()[0];

  var timestamp = rawData[0];
  var area = rawData[1].toString().trim();
  var householdSize = parseInt(rawData[2]);
  var dailyGridHours = rawData[3].toString().trim();
  var altEnergySource = rawData[4].toString().trim();
  var monthlyIncome = rawData[5].toString().trim();
  var energySpending = rawData[6].toString().trim();
  var productiveHours = rawData[7].toString().trim();
  var lifeSatisfaction = rawData[8].toString().trim();
  var biggestChallenge = rawData[9].toString().trim();

  // Clean area name
  area = area.replace(/\w\S*/g, function(txt) {
    return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
  });

  // Clean household size
  if (isNaN(householdSize) || householdSize < 1) householdSize = null;

  // Grid hours to numeric
  var gridHoursNumeric;
  if (dailyGridHours === "0 hours (no grid supply)") gridHoursNumeric = 0;
  else if (dailyGridHours === "1 - 3 hours") gridHoursNumeric = 2;
  else if (dailyGridHours === "4 - 6 hours") gridHoursNumeric = 5;
  else if (dailyGridHours === "7 - 12 hours") gridHoursNumeric = 9.5;
  else if (dailyGridHours === "More than 12 hours") gridHoursNumeric = 15;
  else gridHoursNumeric = null;

  // Income to numeric (bracket midpoints)
  var incomeNumeric;
  if (monthlyIncome === "Below ₦50,000") incomeNumeric = 25000;
  else if (monthlyIncome === "₦50,000 - ₦100,000") incomeNumeric = 75000;
  else if (monthlyIncome === "₦100,001 - ₦200,000") incomeNumeric = 150000;
  else if (monthlyIncome === "₦200,001 - ₦500,000") incomeNumeric = 350000;
  else if (monthlyIncome === "Above ₦500,000") incomeNumeric = 750000;
  else incomeNumeric = null;

  // Spending to numeric (bracket midpoints)
  var spendingNumeric;
  if (energySpending === "Nothing — I don't use alternatives") spendingNumeric = 0;
  else if (energySpending === "Below ₦5,000") spendingNumeric = 2500;
  else if (energySpending === "₦5,000 - ₦15,000") spendingNumeric = 10000;
  else if (energySpending === "₦15,001 - ₦30,000") spendingNumeric = 22500;
  else if (energySpending === "₦30,001 - ₦50,000") spendingNumeric = 40000;
  else if (energySpending === "Above ₦50,000") spendingNumeric = 65000;
  else spendingNumeric = null;

  // Productive hours to numeric
  var productiveHoursNumeric;
  if (productiveHours === "Less than 2 hours") productiveHoursNumeric = 1;
  else if (productiveHours === "2 - 4 hours") productiveHoursNumeric = 3;
  else if (productiveHours === "5 - 7 hours") productiveHoursNumeric = 6;
  else if (productiveHours === "8 - 10 hours") productiveHoursNumeric = 9;
  else if (productiveHours === "More than 10 hours") productiveHoursNumeric = 11;
  else productiveHoursNumeric = null;

  // Extract satisfaction score number
  var satisfactionScore = parseInt(lifeSatisfaction.charAt(0));
  if (isNaN(satisfactionScore)) satisfactionScore = null;

  // Calculate energy burden percentage
  var energyBurdenPct;
  if (incomeNumeric && incomeNumeric > 0 && spendingNumeric !== null) {
    energyBurdenPct = parseFloat(((spendingNumeric / incomeNumeric) * 100).toFixed(2));
  } else {
    energyBurdenPct = null;
  }

  // Write cleaned row to Cleaned Data sheet
  cleanSheet.appendRow([
    timestamp,
    area,
    householdSize,
    dailyGridHours,
    altEnergySource,
    monthlyIncome,
    energySpending,
    productiveHours,
    satisfactionScore,
    biggestChallenge,
    incomeNumeric,
    spendingNumeric,
    gridHoursNumeric,
    productiveHoursNumeric,
    energyBurdenPct
  ]);
}

function createTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
}
