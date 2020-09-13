const fs = require('fs');
const readline = require('readline');
const {google} = require('googleapis');
const {backOff} = require('exponential-backoff');
const peopleJSON = require('./people.json');

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/calendar'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = 'token.json';

const CALENDAR_ID = "cue3vcqj17138sf8u7ai4isv0s@group.calendar.google.com";
const SHEET_ID = "180IRuWqarK_Q4iZE8lNsif0iQfllaGeobqILHbgEtng";
const DATA_RANGE = 'B4:S21'; // SHould be :S21 
const EID_RANGE = 'A1:O17'; // SHould be :O17
const DATA_SHEET_NAME = 'Sheet1';
const EID_SHEET_NAME = 'eids';

const daysJSON = {
    "Monday" : 0,
    "Tuesday" : 1,
    "Wednesday" : 2,
    "Thursday" : 3,
    "Friday" : 4,
    "Saturday" : 5,
    "Sunday" : 6,
}

// Load client secrets from a local file.
fs.readFile('credentials.json', (err, content) => {
    if (err) return console.log('Error loading client secret file:', err);
    // Authorize a client with credentials, then call the Google Sheets API.
    authorize(JSON.parse(content), create);
});

/**
* Create an OAuth2 client with the given credentials, and then execute the
* given callback function.
* @param {Object} credentials The authorization client credentials.
* @param {function} callback The callback to call with the authorized client.
*/
function authorize(credentials, callback) {
    const {client_secret, client_id, redirect_uris} = credentials.installed;
    const oAuth2Client = new google.auth.OAuth2(
        client_id, client_secret, redirect_uris[0]);
        
        // Check if we have previously stored a token.
        fs.readFile(TOKEN_PATH, (err, token) => {
            if (err) return getNewToken(oAuth2Client, callback);
            oAuth2Client.setCredentials(JSON.parse(token));
            callback(oAuth2Client);
        }
    );
}

/**
* Get and store new token after prompting for user authorization, and then
* execute the given callback with the authorized OAuth2 client.
* @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
* @param {getEventsCallback} callback The callback for the authorized client.
*/
function getNewToken(oAuth2Client, callback) {
    const authUrl = oAuth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: SCOPES,
    });
    console.log('Authorize this app by visiting this url:', authUrl);
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout,
    });
    rl.question('Enter the code from that page here: ', (code) => {
        rl.close();
        oAuth2Client.getToken(code, (err, token) => {
            if (err) return console.error('Error while trying to retrieve access token', err);
            oAuth2Client.setCredentials(token);
            // Store the token to disk for later program executions
            fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
                if (err) return console.error(err);
                console.log('Token stored to', TOKEN_PATH);
            });
            callback(oAuth2Client);
        });
    });
}

// Main function called
async function create(auth) {
    const sheets = google.sheets({version: 'v4', auth});
    const calendar = google.calendar({version: 'v3', auth});
    
    try {
        // Get events from the sheet. existingEvents is the ones that already exist (from eids sheet). targetEvents are the ones that are generated from the sheet.
        const existingEvents = await getEvents(sheets, calendar, auth);
        const targetEvents = await generateEventsJSON(sheets);
        const eids = [];
        
        let nCreated = 0;
        let nPatched = 0;
        
        await Promise.all(targetEvents.map(async (row, nRow) => {
            eids[nRow] = [];
            let nPatchedInner = 0;
            let nCreatedInner = 0;
            await Promise.all(row.map(async (targetEvent, nCol) => {
                // If event already exists, check if there are any updates needed
                if (existingEvents[nRow] && existingEvents[nRow][nCol]) {
                    await backoff(() => patchEvent(calendar, auth, existingEvents[nRow][nCol].id, targetEvent));
                    nPatchedInner++;
                } else if (targetEvent) {
                    let response = await backoff(() => insertEvent(calendar, auth, targetEvent));
                    eids[nRow][nCol] = response.data.id;
                    nCreatedInner++;
                }
            }));
            console.log("Finished patching " + nPatchedInner + " events on row " + nRow);
            console.log("Finished creating " + nCreatedInner + " events on row " + nRow);
            nPatched += nPatchedInner;
            nCreated += nCreatedInner;
        }));
        
        console.log("Finished patching " + nPatched + " events");
        console.log("Finished creating " + nCreated + " events");
        
        await backoff(() => patchEids(sheets, auth, eids));
    } catch (e) {
        console.log(e);
    }
    
}

// Exponential backoff function to ensure all requests get through.
function backoff(func) {
    try {
        return backOff(func, {
            jitter: "full",
            startingDelay: Math.floor(Math.random() * 250) + 1000,
            delayFirstAttempt: true,
            numOfAttempts: 11,
            retry : (e, n) => {
                console.log("Failing attempt " + (n - 1));
                return e.code == 403 || e.code >= 500;
            }
        })
    } catch (e) {
        console.log(e);
    }
}

// Get all the existing events of the sheet.
async function getEvents(sheets, calendar, auth) {
    let events = [];
    
    try {
        const response = await getEidSheet(sheets);
        
        const rows = response.data.values;
        if (!rows) {
            return [[]];
        }
        
        let nEvents = 0;
        
        await Promise.all(rows.map(async (col, nRow) => {
            events[nRow] = [];
            let nEventsInner = 0;
            await Promise.all(col.map(async (cell, nCol) => {
                if (rows[nRow][nCol]) {
                    let response = await backoff(() => getEvent(calendar, rows[nRow][nCol]));
                    events[nRow][nCol] = response.data;
                } else {
                    events[nRow][nCol] = null;
                }
                nEventsInner++;
            }));
            console.log("Finished getting " + nEventsInner + " events on row " + nRow);
            nEvents += nEventsInner;
        }));
        console.log("In total got " + nEvents + " events");
        
    } catch (e) {
        console.log(e);
    }
    
    return events;
    
}

// Generate list of all event JSON.
async function generateEventsJSON(sheets) {
    let events = [];
    
    try {
        const response = await getDataSheet(sheets);
        
        const rows = response.data.values;
        const nRows = rows.length;
        
        for (let nRow = 1; nRow < nRows; nRow++) {
            events[nRow - 1] = [];
            for (let nCol = 3; nCol < rows[nRow].length; nCol++) {
                if (rows[nRow][nCol]) {
                    events[nRow - 1][nCol - 3] = generateEventJSON(rows, nRow, nCol);
                }
            }
        }
        
    } catch (e) {
        console.log(e);
    }
    
    return events;
    
}

// From the sheet figure out what a row, col event would look like
function generateEventJSON(data, nRow, nCol) {
    const cell = data[nRow][nCol];
    let time = getTime(data[nRow][2], data[0][nCol]);
    let people = getPeople(cell);
    return {
        "summary": data[nRow][0],
        "description": data[nRow][1],
        "attendees" : people, 
        'start': {
            'dateTime': time.toISOString(),
            'timeZone': 'Australia/Sydney'
        },
        'end': {
            'dateTime': addHours(time, 1).toISOString(),
            'timeZone': 'Australia/Sydney'
        },
        'reminders': {
            'useDefault': false,
            'overrides': [
                {'method': 'email', 'minutes': 24 * 60},
                {'method': 'popup', 'minutes': 10}
            ]
        }
    };
}

// Async Requests

function getEvent(calendar, eid) {
    return calendar.events.get({
        "calendarId": CALENDAR_ID,
        "eventId": eid
    });
}

function getDataSheet(sheets) {
    return sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: DATA_SHEET_NAME + "!" + DATA_RANGE,
    });
}

function insertEvent(calendar, auth, eventJSON) {
    return calendar.events.insert({
        "auth": auth,
        'calendarId': CALENDAR_ID,
        'resource' : eventJSON
    });
}

function patchEvent(calendar, auth, eid, patchJSON) {
    return calendar.events.patch({
        "auth": auth,
        'calendarId': CALENDAR_ID,
        'eventId': eid,
        'resource' : patchJSON
    });
}

function patchEids(sheets, auth, values) {
    return sheets.spreadsheets.values.update({
        "spreadsheetId": SHEET_ID,
        "range": EID_SHEET_NAME + '!' + EID_RANGE,
        "auth": auth,
        'valueInputOption': "RAW",
        'resource': {values}
    })
}

function getEidSheet(sheets) {
    return sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: EID_SHEET_NAME + "!" + EID_RANGE,
    });
}

// Random util functions

// Given a string of form 'HS EL' etc (refer to peopleJSON) converts into a list of emails and display names
function getPeople(string) {
    let names = string.split(' ');
    return names.map((name) => {return {"email": peopleJSON[name].email, "displayName": peopleJSON[name].name};});
}

// Given a date string in the form dd/MM/yyyy and a number of days to add to it returns a date object
function getTime(daysToAdd, date) {
    let dateArray = date.split("/");
    let day = dateArray[0];
    let month = dateArray[1] - 1;
    let year = dateArray[2];
    let datetime = new Date(year, month, day, 9, 0, 0, 0);
    datetime = addDays(datetime, daysJSON[daysToAdd]);
    return datetime;
}

// From https://stackoverflow.com/questions/563406/add-days-to-javascript-date
// Adds days many days to date and returns it
function addDays(date, days) {
    var result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
}

// Adds hours many hours to the given date
function addHours(date, hours) {
    var result = new Date(date);
    result.setHours(result.getHours() + hours);
    return result;
}
