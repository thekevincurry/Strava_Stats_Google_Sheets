function refreshToken() {
  var scriptProperties = PropertiesService.getScriptProperties();
  
  // Retrieve your Strava app's client ID, client secret, and the current refresh token from script properties
  var clientId = scriptProperties.getProperty('STRAVA_CLIENT_ID');
  var clientSecret = scriptProperties.getProperty('STRAVA_CLIENT_SECRET');
  var refreshToken = scriptProperties.getProperty('STRAVA_REFRESH_TOKEN');

  var tokenEndpoint = 'https://www.strava.com/oauth/token';
  var payload = {
    'client_id': clientId,
    'client_secret': clientSecret,
    'grant_type': 'refresh_token',
    'refresh_token': refreshToken
  };

  var options = {
    'method' : 'post',
    'payload' : payload,
    'muteHttpExceptions': true // to better handle HTTP exceptions
  };

  var response = UrlFetchApp.fetch(tokenEndpoint, options);
  
  // It's a good practice to handle HTTP errors gracefully
  if (response.getResponseCode() != 200) {
    Logger.log('Error refreshing token: ' + response.getContentText());
    return null;
  }

  var json = JSON.parse(response.getContentText());

  // Update the stored access and refresh tokens
  scriptProperties.setProperty('STRAVA_ACCESS_TOKEN', json.access_token);
  scriptProperties.setProperty('STRAVA_REFRESH_TOKEN', json.refresh_token);

  return json.access_token;
}


function getStravaStats() {
  const accessToken = refreshToken(); // Refresh the token
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const url = 'https://www.strava.com/api/v3/athletes/22032/stats';
  const headers = {
    "Authorization": "Bearer " + accessToken
  };

  const options = {
    "method" : "get",
    "headers" : headers,
    "muteHttpExceptions": true
  };

  // Fetch the data from Strava API
  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  // Assuming you want to get the year to date ride totals
  const ytdRideTotals = json.ytd_ride_totals;

  // Define the row in the Google Sheet where the data starts
  const startRow = 2;

  // Update the Google Sheet with the fetched data
  sheet.getRange(startRow, 1).setValue("YTD Ride Distance (meters)");
  sheet.getRange(startRow, 2).setValue(ytdRideTotals.distance);
  sheet.getRange(startRow + 1, 1).setValue("YTD Ride Time (seconds)");
  sheet.getRange(startRow + 1, 2).setValue(ytdRideTotals.moving_time);
  // ...add more fields as necessary
}

// Add a custom menu to run the script from the Google Sheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Strava Stats')
      .addItem('Fetch YTD Stats', 'getStravaStats')
      .addToUi();
}
