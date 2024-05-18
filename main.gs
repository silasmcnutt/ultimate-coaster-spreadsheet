function onChange(e) {
  try {
    // Check if the event object exists and contains the expected properties
    if (e.changeType == 'EDIT' || e.changeType == 'FORMAT' || e.changeType == 'OTHER') {
      let range = e.source.getActiveRange();
      let column = range.getColumn();
      let row = range.getRow();
      let sheetData = range.getSheet();
      let sheetName = sheetData.getName();

      // DATA INPUT SHEET
      if (sheetName == 'Data Input') {
        if (column == 1 || column == 2) {
          if (row >= 7) {
            let ride_name = sheetData.getRange(row, 1).getValue();
            let park_name = sheetData.getRange(row, 2).getValue();

            Logger.log('Ride: ' + ride_name);
            Logger.log('Park: ' + park_name)
            if (ride_name != '' || ride_name != null || park_name != '' || park_name != null) {
              sheetData.getRange(row, 3).setValue(getIdByNameAndPark(ride_name, park_name));
            }
          }
        } else if (column == 3) {
          if (row >= 7) {
            let id = sheetData.getRange(row,3).getValue()

            Logger.log('Ride ID: ' + id)
            let data = getRideData(id)
            sheetData.getRange(row,1).setValue(data['name'])
            sheetData.getRange(row,2).setValue(data['park']['name'])
          }
        }
      }
    } else {
      Logger.log('Event object is undefined or does not contain the expected properties.');
    }
  } catch (error) {
    Logger.log("Error: " + error);
  }
}

function primaryStatsRefresh() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getSheetByName("Data Input");
  var targetSheet = spreadsheet.getActiveSheet();

  // Copy data from source column C to target column A
  var sourceValues = sourceSheet.getRange("C6:C").getValues();
  targetSheet.getRange(3, 1, sourceValues.length, 1).setValues(sourceValues);

  // Get the sheet where you want to iterate through the column
  var sheetData = spreadsheet.getSheetByName("Primary Stats");
  
  // Get the range of values in the desired column (column A in this example)
  let columnValues = sheetData.getRange("A:A").getValues().filter(row => row[0] !== '');
  Logger.log(columnValues)
  
  // Iterate through each value in the column
  for (var i = 2; i < columnValues.length; i++) {
    var value = columnValues[i][0]; // Get the value from the current row
    
    // Perform actions based on the value
    if (value !== null && value !== '') { // Check if the cell is not empty
      Logger.log("Row " + (i+1) + ": Coaster ID " + value); // Log the value to the console
      var data = getRideData(value);
      var park_name = data['park']['name'];
      var ride_name = data['name'];
      var state = data['state'];
      var make = data['make'];
      var type = data['type'];
      var status = data['status']['state'];
        
      // Set values in the same row as the edited ID
      sheetData.getRange(i+2, 2).setValue(ride_name); // Set park_name in column B of the same row
      sheetData.getRange(i+2, 3).setValue(park_name); // Set ride_name in column C of the same row
      sheetData.getRange(i+2, 4).setValue(state);
      sheetData.getRange(i+2, 5).setValue(make);
      sheetData.getRange(i+2, 6).setValue(type);
      sheetData.getRange(i+2, 7).setValue(status);
      SpreadsheetApp.flush();
    }
  }
}

function extendedStatsRefresh() {
  // Define spreadsheets
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let dataInput = spreadsheet.getSheetByName("Data Input");
  let extendedStats = spreadsheet.getActiveSheet();

  // Add IDs to Extended Stats sheet
  let sourceIDs = dataInput.getRange("C6:C").getValues();
  extendedStats.getRange(3, 1, sourceIDs.length, 1).setValues(sourceIDs);

  let columnValues = extendedStats.getRange("A:A").getValues().filter(row => row[0] !== '');
  let units = dataInput.getRange("H12").getValues();

  for (let i = 2; i < columnValues.length; i++) {
    let id = columnValues[i][0];

    if (id !== null && id !== '') {
      Logger.log("Row " + (i+1) + ": Coaster ID " + id);
      let data = getRideData(id);
      let park_name = data['park']['name'];
      let ride_name = data['name'];
      let state = data['state'];
      let make = data['make'];
      let type = data['type'];
      let design = data['design']
      let length = data['stats']['length']
      let height = data['stats']['height']
      let drop = data['stats']['drop']
      let speed = data['stats']['speed']
      let duration = data['stats']['duration']
      if (duration != '' && duration != null) {
        duration = convertTime(duration)
      }
      let inversions = data['stats']['inversions']
      let status = data['status']['state']
      let opened = data['status']['date']['opened']
      let closed = data['status']['date']['closed']
      let cost = data['stats']['cost']
      Logger.log("Cost: " + cost)
      if (cost === undefined) {
        Logger.log('Querying for secondary cost...')
        cost = secondaryCostQuery(ride_name, park_name)
      }

      if (units == 'Imperial') {
        extendedStats.getRange("H3").setValue('Length (ft)')
        extendedStats.getRange("I3").setValue('Height (ft)')
        extendedStats.getRange("J3").setValue('Drop (ft)')
        extendedStats.getRange("K3").setValue('Speed (mph)')

        if (!Array.isArray(length) && length != '' && length != null) {
          length = Math.round(length * 3.281);
        }
        if (!Array.isArray(height) && height != '' && height != null) {
          height = Math.round(height * 3.281);
        }
        if (!Array.isArray(drop) && drop != '' && drop != null) {
          drop = Math.round(drop * 3.281);
        }
        if (!Array.isArray(speed) && speed != '' && speed != null) {
          speed = Math.round(speed/1.609);
        }
      }
      else if (units == 'Metric') {
        extendedStats.getRange("H3").setValue('Length (m)')
        extendedStats.getRange("I3").setValue('Height (m)')
        extendedStats.getRange("J3").setValue('Drop (m)')
        extendedStats.getRange("K3").setValue('Speed (kph)')     
      }

      
        
      // Place data
      extendedStats.getRange(i+2, 2).setValue(ride_name);
      extendedStats.getRange(i+2, 3).setValue(park_name);
      extendedStats.getRange(i+2, 4).setValue(state);
      extendedStats.getRange(i+2, 5).setValue(make);
      extendedStats.getRange(i+2, 6).setValue(type);
      extendedStats.getRange(i+2, 7).setValue(design);
      extendedStats.getRange(i+2, 8).setValue(length);
      extendedStats.getRange(i+2, 9).setValue(height);
      extendedStats.getRange(i+2, 10).setValue(drop);
      extendedStats.getRange(i+2, 11).setValue(speed);
      extendedStats.getRange(i+2, 12).setValue(duration);
      extendedStats.getRange(i+2, 13).setValue(inversions);
      extendedStats.getRange(i+2, 14).setValue(status);
      extendedStats.getRange(i+2, 15).setValue(opened);
      extendedStats.getRange(i+2, 16).setValue(closed);
      extendedStats.getRange(i+2, 17).setValue(cost)
      SpreadsheetApp.flush();
    }
  }
}

function primaryStatsClearData() {
  // Define spreadsheets
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let primaryStats = spreadsheet.getActiveSheet();
  primaryStats.getRange('A4:G').clearContent()
}

function extendedStatsClearData() {
  // Define spreadsheets
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let extendedStats = spreadsheet.getActiveSheet();
  extendedStats.getRange('A4:Q').clearContent()
}



function getIdByNameAndPark(coaster_name, park_name) {
  try {
    // Initial API call for coasters
    Logger.log('Fetching all coasters from RCDB...');
    var response = UrlFetchApp.fetch("https://rcdb-api.vercel.app/api/coasters/search?q=" + coaster_name);
    var json = response.getContentText();
    var data = JSON.parse(json);
    var coasters = [];
    var parks = [];

    // Add all possible coasters to array
    for (var i = 0; i < data['coasters'].length; i++) {
      coasters.push(data['coasters'][i]['id']);
      parks.push(data['coasters'][i]['park']['id']);
    }
    Logger.log('Coasters: ' + coasters);
    Logger.log('Parks: ' + parks);

    // Secondary API call for Parkname if multiple coasters found
    if (data['totalMatch'] != 1) {
      Logger.log('Searching for park from RCDB...');
      var response = UrlFetchApp.fetch("https://rcdb-api.vercel.app/api/coasters/search?q=" + park_name);
      var json = response.getContentText();
      var data = JSON.parse(json);
      var confirmed_coaster;
      var park = data['coasters'][0]['park']['id'];

      for (var i = 0; i < parks.length; i++) {
        if (park == parks[i]) {
          confirmed_coaster = coasters[i];
          Logger.log('Confirmed coaster ID: ' + confirmed_coaster);
          break;
        }
      }
    } else if (data['totalMatch'] == 1) {
      confirmed_coaster = coasters[0]
    }

    // If a valid coaster ID is found, return it
    if (confirmed_coaster) {
      return confirmed_coaster;
    }
  } catch (error) {
    Logger.log("Error: " + error);
  }

  // Return null if no valid coaster ID is found
  return null;
}


function getRideData(id) {
  var maxRetries = 3; // Maximum number of retry attempts
  var retryCount = 0;
  
  while (retryCount < maxRetries) {
    try {
      var response = UrlFetchApp.fetch("https://rcdb-api.vercel.app/api/coasters/" + id);
      var json = response.getContentText();
      var data = JSON.parse(json);
      
      // If fetching data is successful, return it
      return data;
    } catch (error) {
      // If an error occurs, log it and retry
      Logger.log("Error fetching ride data. Retrying...");
      retryCount++;
    }
  }
  
  // If max retries reached and still unsuccessful, log an error and return null
  Logger.log("Max retry attempts reached. Unable to fetch ride data.");
  return null;
}

function convertTime(minutesSeconds) {
  // Assuming the input format is already in "number:number" (minutes:seconds)
  var parts = minutesSeconds.toString().split(":");
  var minutes = parseInt(parts[0]);
  var seconds = parseInt(parts[1]);
  
  // Convert minutes and seconds to total seconds
  var totalSeconds = minutes * 60 + seconds;
  
  // Calculate hours, minutes, and seconds
  var hours = Math.floor(totalSeconds / 3600);
  var remainingSeconds = totalSeconds % 3600;
  var formattedMinutes = Math.floor(remainingSeconds / 60);
  var formattedSeconds = remainingSeconds % 60;
  
  // Ensure the minutes and seconds are formatted with leading zero if needed
  if (formattedMinutes < 10) {
    formattedMinutes = "0" + formattedMinutes;
  }
  if (formattedSeconds < 10) {
    formattedSeconds = "0" + formattedSeconds;
  }
  
  return hours + ":" + formattedMinutes + ":" + formattedSeconds;
}

function secondaryCostQuery(rideName, parkName) {
  var maxRetries = 3;
  var retryCount = 0;
  
  // Find correct article name
  while (retryCount < maxRetries) {
    try {
      var response = UrlFetchApp.fetch("https://en.wikipedia.org/w/api.php?action=query&list=search&srsearch=" + rideName + " " + parkName + "&format=json");
      var json = response.getContentText();
      var data = JSON.parse(json);
      var confirmedArticleName = data['query']['search'][0]['title'];
      Logger.log(confirmedArticleName)
      // If fetching data is successful, return it
      break;
    } catch (error) {
      Logger.log("Error fetching article name. Retrying...");
      retryCount++;
    }
    Logger.log("Max retry attempts reached. Unable to fetch article name.");
    return null;
  }
  

  var url = 'https://en.wikipedia.org/wiki/' + confirmedArticleName;

  try {
    var response = UrlFetchApp.fetch(url);
    var html = response.getContentText();
    var $ = Cheerio.load(html);
    var infoboxTable = $('.infobox');
    var costRow = infoboxTable.find('th:contains("Cost")').parent();
    var costValue = costRow.find('td').text().trim();
    
    costValue = convertToNumber(costValue);
    Logger.log(costValue)
    return costValue;
  } catch (error) {
    Logger.log("Error fetching cost from Wikipedia: " + error);
    return null;
  }
}

function convertToNumber(costString) {
  costString = costString.toLowerCase();

  // Remove any characters until it reaches a number
  var numericString = costString.replace(/^[^\d]*/, '');
  Logger.log(numericString)

  // Check if the string contains '[' character
  var indexOfBracket = numericString.indexOf('[');
  if (indexOfBracket !== -1) {
    // Truncate the string including ']' and characters after it
    numericString = numericString.slice(0, indexOfBracket);
  }
  Logger.log(numericString)

  // Remove any text that is not the word "million"
  numericString = numericString.replace(/[^a-z\d.]/gi, '');

  // Check if the string contains the word "million"
  if (numericString.includes('million')) {
    // Remove the word "million" and multiply the number by 1 million
    numericString = numericString.replace('million', '');
    var numericValue = parseFloat(numericString) * 1000000;
    return numericValue;
  } else {
    // If "million" is not found, return null or any default value you prefer
    return numericString;
  }
}




// Article Query https://en.wikipedia.org/w/api.php?action=query&list=search&srsearch=Ride Name Park Name&format=json
// Complete Info Query https://en.wikipedia.org//w/api.php?action=query&format=json&prop=extracts&titles=Ride Name Here&formatversion=2&rvprop=content&rvslots=*

