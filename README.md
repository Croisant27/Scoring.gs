# Scoring.gs
 'Scoring.gs' file - Data manipulation and delivering outcome

function generateVariablesFromSheet() {
  // Get reference to the active spreadsheet and the "Form Responses 1" sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = spreadsheet.getSheetByName("Form Responses 1");

  // Initialize an object to store generator data
  var generators = {};

  // Get the data from the sheet
  var data = dataSheet.getDataRange().getValues();

  // Iterate through the rows of data (excluding the header)
  for (var i = 1; i < data.length; i++) {
    // Extract values from the current row
    var generatorType = data[i][1];
    var groupSellingPowerFirst = data[i][2];
    var groupSellingPowerCompound = data[i][5];
    var price1 = data[i][3];
    var price2 = data[i][6];

    // Generate keys for each generator type 
    var key = generatorType.replace(/\s+/g, '');

    // If the generator type does not exist in the object, initialize it
    if (!generators[key]) {
      generators[key] = {
        price1: [],
        price2: [],
        powerFirst: [],
        powerCompound: []
      };
    }

    // Append the values to the arrays in the object
    generators[key].price1.push(price1);
    generators[key].price2.push(price2);
    generators[key].powerFirst.push(groupSellingPowerFirst);
    generators[key].powerCompound.push(groupSellingPowerCompound);
  }

  // Log the structured data for monitoring
  console.log(generators);
  return generators;
}

// Define the function to transform and flatten generator data
function flattenGeneratorData(generators) {
  // Initialize an empty array to hold all offer objects
  let allOffers = [];

  // Iterate over each key (generator type) in the 'generators' object
  Object.keys(generators).forEach(generatorType => {
    // Retrieve the specific generator object for the current iteration
    let generator = generators[generatorType];
    
    // Iterate over each price in the generator's 'price1' array
    generator.price1.forEach((price, index) => {
      // Check if the price is greater than 0 to ensure it's a valid offer
      if (price > 0) {
        // Push a new object into the 'allOffers' array representing this offer
        allOffers.push({
          type: generatorType, // The generator type (e.g., Coal, GasA)
          price: price, // The price for this offer
          power: generator.powerFirst[index], // The amount of power offered at this price
          bidType: 'First Bid', 
        });
      }
    });
    
    // Similar to the above, iterate over each price in the generator's 'price2' array
    generator.price2.forEach((price, index) => {
      // Again, check for valid offers with price greater than 0
      if (price > 0) {
        // Push a new offer object into 'allOffers'
        allOffers.push({
          type: generatorType, // The generator type
          price: price, // The price for this offer
          power: generator.powerCompound[index], // The amount of power offered at this price
          bidType: 'Compound Bid', 
        });
      }
    });
  });

  // Return the array of all offers, now containing a flat list of all offers
  // from all generators, each represented as an object with details about the type,
  // price, power, and whether it's for 'First Bid' or 'Compound Bid'
  return allOffers;
}

// Define the function to sort offer objects by their price
function sortOffersByPrice(offers) {
  return offers.sort((a, b) => a.price - b.price);
}

// Define the demand for each round
const roundDemands = {
  r1: 125000,
  r2: 220000,
  r3: 205000,
  r4: 180000,
  r5: 75000
};

function processGeneratorData() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var marketPriceSheet = ss.getSheetByName("strikePrice");
  var formResponsesSheet = ss.getSheetByName("Form responses 1");

  // Read the market price and round information
  var marketPrice = parseFloat(marketPriceSheet.getRange("G1").getValue());
  var roundInfo = formResponsesSheet.getRange("I2").getValue();
  var demand = roundDemands[roundInfo]; // Get demand for the current round

  // Initialize variables to track the cumulative total power sold
  var cumulativeTotalPowerSold = 0;
  var previousPricePointTotalPower = 0;

  // Obtain the 'generators' object 
  var generators = generateVariablesFromSheet(); // Ensure this is adapted for your actual data retrieval

  // Flatten and sort the generator data by price
  var flattenedData = flattenGeneratorData(generators);
  var sortedOffers = sortOffersByPrice(flattenedData);

  // Unit costs for each generator type
  var unitCosts = {
    Coal: 30,
    GasA: 50, // Assuming "Gas A" becomes "GasA" in your key naming
    GasB: 50,
    GasC: 50,
    Hydro: 0,
    PumpedStorage: 0,
    OnshoreWind: 0,
    OffshoreWind: 0,
    SolarPV: 0,
    Nuclear: 20
  };

  // Reference or create a sheet for output

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheetName = "Sorted Offers " + roundInfo;
  var outputSheet = ss.getSheetByName(outputSheetName);

  // Check if the sheet exists, if not, create it. If it exist,  clear it.
  if (!outputSheet) {
    outputSheet = ss.insertSheet(outputSheetName);
  } else {
    outputSheet.clear(); // Clear the existing data if running the same round again
  }

  // Set headers for the table
  outputSheet.appendRow(["Generator Type", "Bid Type", "Price (£/MWh)", "Amount to Supply (MWh)", "Unit Cost (£/Mwh)", "Shortfall (MWh)", "Multiplier", "Energy Sold (MWh)", "Profit (£)"]);

  // Get priceRange data for shortfall and multiplier calculations
  var priceRange = generatePriceRangeFromSheet();

  // Populate the table with sorted offers
  // This checks if the bidType property of the current offer object is equal to 'First Bid'or 'Compound Bid'.
  sortedOffers.forEach((offer, index) => {
    var bidType = offer.bidType === 'First Bid' ? 'First Bid' : 'Compound Bid';

    var unitCost = unitCosts[offer.type] || 0; // Retrieve the unit cost, defaulting to 0 if not found
    var totalPowerSoldUntilPreviousPricePoint = priceRange[offer.price] ? priceRange[offer.price].PreviousTotal[0] : 0; // Retrieve total power sold until the previous price point
    var shortfall = Math.max(0, demand - totalPowerSoldUntilPreviousPricePoint); // Calculate shortfall, ensuring it doesn't go below 0
    var powerSoldAtCurrentPrice = priceRange[offer.price] ? priceRange[offer.price].Customers[0] : offer.power; // Retrieve power sold at current price point
    var multiplier = Math.min(1, shortfall / powerSoldAtCurrentPrice); // Calculate multiplier

    var energySoldFinal = offer.power * multiplier;

    var profit = offer.price < marketPrice ? energySoldFinal * (offer.price - unitCost) : energySoldFinal * (marketPrice - unitCost);
    
    // Append row with new column data
    outputSheet.appendRow([
      offer.type,
      offer.bidType,
      offer.price,
      offer.power,
      unitCost,
      shortfall,
      multiplier,
      energySoldFinal,
      profit
    ]);
  });

  //%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
  // After appending all rows, set number formatting for the entire column to ensure commas are used
var lastRow = outputSheet.getLastRow();
if (lastRow > 1) { // Check if there are any data rows
  var numRows = lastRow - 1; // Calculate the number of data rows
  // Update column letters as needed based on your actual spreadsheet structure
  var priceRange = outputSheet.getRange(2, 3, numRows); // Column C: Price (£/MWh)
  var powerRange = outputSheet.getRange(2, 4, numRows); // Column D: Amount to Supply (MWh)
  var unitCostRange = outputSheet.getRange(2, 5, numRows); // Column E: Unit Cost (£/MWh)
  var shortfallRange = outputSheet.getRange(2, 6, numRows); // Column F: Shortfall (MWh)
  var multiplierRange = outputSheet.getRange(2, 7, numRows); // Column G: Multiplier
  var energySoldRange = outputSheet.getRange(2, 8, numRows); // Column H: Energy Sold (MWh)
  var profitRange = outputSheet.getRange(2, 9, numRows); // Column I: Profit (£)

  // Apply number formatting with commas for thousands separator
  var numberFormat = "#,##0.00";
  priceRange.setNumberFormat(numberFormat);
  powerRange.setNumberFormat(numberFormat);
  unitCostRange.setNumberFormat(numberFormat);
  shortfallRange.setNumberFormat(numberFormat);
  multiplierRange.setNumberFormat("0.000"); // Multiplier visualized in 3 decimal places
  energySoldRange.setNumberFormat(numberFormat);
  profitRange.setNumberFormat("#,##0"); // Profit as an integer (no decimal)
}

  //%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
}

function generatePriceRangeFromSheet() {
  // Get reference to the active spreadsheet and the "Sheet1" sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = spreadsheet.getSheetByName("strikePrice");

  // Initialize an object to store price range data
  var priceRange = {};

  // Get the data from the sheet, focusing on columns A, B, and C
  var data = dataSheet.getRange("A:C").getValues();

  // Iterate through the rows of data (excluding the header)
  for (var i = 1; i < data.length; i++) {
  // Skip the loop if no more data is available or if the first cell is empty/undefined
  if (data[i][0] === '' || data[i][0] == null) {
    break;
  }

  var price = data[i][0];
  var customers = data[i][1];
  var totalPowerSold = data[i][2];

  // Use price as a key
  var key = price.toString();

  // Initialize the object at this key if it does not exist, including PreviousTotal
  if (!priceRange[key]) {
    priceRange[key] = {
      Customers: [],
      TOTAL: [],
      PreviousTotal: [] // Initialize PreviousTotal array
    };
  }

  // Calculate the total power sold until the previous price point
  // Ensure i > 1 to avoid referencing the header row and initialize for the first data row
  var totalPowerSoldUntilPreviousPricePoint = i > 1 ? parseFloat(data[i-1][2]) : 0;

  // Append current values to the arrays
  priceRange[key].Customers.push(customers);
  priceRange[key].TOTAL.push(totalPowerSold);
  priceRange[key].PreviousTotal.push(totalPowerSoldUntilPreviousPricePoint); // Append the calculated value
}

  // Log the structured data for monitoring
  console.log(priceRange);

  // Return the priceRange object for further use
  return priceRange;
}

function compileGeneratorProfits() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Define the generator types for column headers
  var generatorTypes = ["Coal", "GasA", "GasB", "GasC", "Hydro", "PumpedStorage", "OnshoreWind", "OffshoreWind", "Nuclear", "SolarPV"];
  
  // Create or get the Scoring sheet
  var scoringSheet = ss.getSheetByName("Scoring");
  
  if (!scoringSheet) {
    scoringSheet = ss.insertSheet("Scoring");
  } else {
    // Calculate the last row with content to ensure we don't clear unused rows
    var lastRow = scoringSheet.getLastRow();
    
    scoringSheet.clear(); // Clear the existing data if running again
  }
  
  // Set up headers in the Scoring sheet
  var headers = ["Round"].concat(generatorTypes.map(function(type) { return type + " (£)"; }));
  scoringSheet.appendRow(headers);

   // Styles for the profit table header
  var headerRange = scoringSheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4a86e8');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  
  // Variable to keep track of total profits & power sold for each generator type across all rounds
  var totalProfits = Array(generatorTypes.length).fill(0);
  var totalPowerSold = Array(generatorTypes.length).fill(0);
  
  // Process each round
  for (var round = 1; round <= 5; round++) {
    var roundSheetName = "Sorted Offers r" + round;
    var roundSheet = ss.getSheetByName(roundSheetName);
    
    if (!roundSheet) {
      Logger.log("Sheet for round " + round + " does not exist.");
      continue;
    }
    
    var dataRange = roundSheet.getRange("A2:I" + roundSheet.getLastRow());
    var roundData = dataRange.getValues();
    
    var profitDataForRound = generatorTypes.map(() => 0);
    var powerDataForRound = generatorTypes.map(() => 0);
    
    // Process each row in the round sheet for profits and power sold
    for (var i = 0; i < roundData.length; i++) {
      var row = roundData[i];
      var generatorType = row[0]; // Generator type in column A
      var profit = parseFloat(row[8]); // Profit in column I
      var powerSold = parseFloat(row[7]); // Power sold in column H
      var generatorIndex = generatorTypes.indexOf(generatorType);
      
      if (generatorIndex !== -1) {
        profitDataForRound[generatorIndex] += profit;
        totalProfits[generatorIndex] += profit;
        powerDataForRound[generatorIndex] += powerSold;
        totalPowerSold[generatorIndex] += powerSold;
      }
    }
    
    // Add this round's data to the Scoring sheet
    scoringSheet.appendRow([round].concat(profitDataForRound));
  }

  var lastRow = scoringSheet.getLastRow() + 4; // Add some space between tables
  
  
  // Add TOTAL row
  scoringSheet.appendRow(["TOTAL"].concat(totalProfits));
  
  // Apply number formatting for better readability
  scoringSheet.getRange("B2:K7").setNumberFormat("#,##0");

 // Set up headers in the Scoring sheet
  var powerHeaders = ["Round"].concat(generatorTypes.map(function(type) { return type + " (MWh)"; }));

   // Styles for the profit table header
  var powerHeaderRange = scoringSheet.getRange(lastRow, 1, 1, powerHeaders.length).setValues([powerHeaders]);
  powerHeaderRange.setBackground('#ff9900'); // A different color for distinction
  powerHeaderRange.setFontColor('#ffffff');
  powerHeaderRange.setFontWeight('bold');
  
  // Variable to keep track of total profits & power sold for each generator type across all rounds
  var totalProfits = Array(generatorTypes.length).fill(0);
  var totalPowerSold = Array(generatorTypes.length).fill(0);
  
  // Process each round
  for (var round = 1; round <= 5; round++) {
    var roundSheetName = "Sorted Offers r" + round;
    var roundSheet = ss.getSheetByName(roundSheetName);
    
    if (!roundSheet) {
      Logger.log("Sheet for round " + round + " does not exist.");
      continue;
    }
    
    var dataRange = roundSheet.getRange("A2:I" + roundSheet.getLastRow());
    var roundData = dataRange.getValues();
    
    var profitDataForRound = generatorTypes.map(() => 0);
    var powerDataForRound = generatorTypes.map(() => 0);
    
    // Process each row in the round sheet for profits 
    for (var i = 0; i < roundData.length; i++) {
      var row = roundData[i];
      var generatorType = row[0]; // Generator type in column A
      var profit = parseFloat(row[8]); // Profit in column I
      var powerSold = parseFloat(row[7]); // Power sold in column H
      var generatorIndex = generatorTypes.indexOf(generatorType);
      
      if (generatorIndex !== -1) {
        profitDataForRound[generatorIndex] += profit;
        totalProfits[generatorIndex] += profit;
        powerDataForRound[generatorIndex] += powerSold;
        totalPowerSold[generatorIndex] += powerSold;
      }
    }
    
    // Add this round's data to the Scoring sheet
    scoringSheet.appendRow([round].concat(powerDataForRound));
  }
  
  // Style the power sold table header
  var powerHeaderRange = scoringSheet.getRange(lastRow, 1, 1, powerHeaders.length);
  powerHeaderRange.setBackground('#ff9900'); // A different color for distinction
  powerHeaderRange.setFontColor('#ffffff');
  powerHeaderRange.setFontWeight('bold');

  // Append TOTAL row for power sold
  scoringSheet.appendRow(["TOTAL"].concat(totalPowerSold));
  scoringSheet.getRange("B" + (lastRow + 1) + ":K" + (lastRow + 6)).setNumberFormat("#,##0");
  
  // Apply number formatting for better readability for both tables
  scoringSheet.getRange(2, 2, lastRow + 1 - 2, generatorTypes.length + 1).setNumberFormat("#,##0");

  //%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
  // Replace 'your_form_id_here' with your actual Form ID
  var formId = '1BKNH23_HvfBZxLynFM85fKXA3WCCtpmxgkdtB_IWTCs';
  var form = FormApp.openById(formId);
  
  // Retrieve all form responses and delete them
  var responses = form.getResponses();
  for (var i = 0; i < responses.length; i++) {
    form.deleteResponse(responses[i].getId());
  }
  
  // Optionally, log that form responses have been cleared
  Logger.log('All form responses have been cleared.');

  // Now, clear the spreadsheet responses
  // Replace 'your_spreadsheet_id_here' with your actual Spreadsheet ID
  var spreadsheetId = '190MP6vHsKf5llwZJYw_pBp-HtA6kok_IwW19QRJSSqc';
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheets()[0]; // Assumes responses are in the first sheet
  
  // Clears all data from the sheet, except for headers (assuming row 1 is headers)
  sheet.getRange('A2:' + sheet.getLastColumn() + sheet.getLastRow()).clearContent();
  
  // Optionally, log that spreadsheet responses have been cleared
  Logger.log('All spreadsheet responses have been cleared.');

  var chartRange = scoringSheet.getRange("A2:" + String.fromCharCode(65 + generatorTypes.length) + (6 + 1)); // Adjust range accordingly
  var chartBuilder = scoringSheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(chartRange)
    .setPosition(1, 1, 0, 0)
    .setOption('title', 'Profits by Generator Type Over Rounds')
    .setOption('hAxis', {title: 'Generator Type', minValue: 0})
    .setOption('vAxis', {title: 'Rounds'})
    .setOption('legend', {position: 'right', textStyle: {color: 'blue', fontSize: 12}})  // Explicitly configure legend
    .setOption('isStacked', true)  // Optional: Set to true for a stacked bar chart
    .build();
  scoringSheet.insertChart(chartBuilder);
}


