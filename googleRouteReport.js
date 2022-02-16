//Production folders on shared drive
const CHANNEL_LAUNCH = "INSERT_FOLDER_ID"; //Channel launch folder
const DEPARTMENT_FOLDER = "INSERT_FOLDER_ID"; //BE and Comp support folder
const DAILY_LOGS_FOLDER = "INSERT_FOLDER_ID"; //Daily Router Logs folder
const PATH_VERIFICATION = "INSERT_FOLDER_ID"; //Path Verifications folder
const SIGNAL_INTEGRITY = "INSERT_FOLDER_ID"; //Signal Integrity routes folder
const TAKEDOWN_RESTORE = "INSERT_FOLDER_ID"; //Takedown/Restore folder
const WEEKLY_REPORT_FOLDER = DriveApp.getFolderById("INSERT_FOLDER_ID"); //Weekly report folder
const FILE = "INSERT_FOLDER_ID"; //ID of the template file

////Test folders on my drive
//const CHANNEL_LAUNCH = "INSERT_FOLDER_ID"; //Channel launch folder
//const DEPARTMENT_FOLDER = "INSERT_FOLDER_ID"; //BE and Comp support folder
//const DAILY_LOGS_FOLDER = "INSERT_FOLDER_ID"; //Daily Router Logs folder
//const PATH_VERIFICATION = "INSERT_FOLDER_ID"; //Path Verifications folder
//const SIGNAL_INTEGRITY = "INSERT_FOLDER_ID"; //Signal Integrity routes folder
//const TAKEDOWN_RESTORE = "INSERT_FOLDER_ID"; //Takedown/Restore folder
//const WEEKLY_REPORT_FOLDER = DriveApp.getFolderById("INSERT_FOLDER_ID"); //Weekly report folder
//const FILE = "INSERT_FOLDER_ID"; //ID of the template file

let date = new Date();
let channelLaunchArray = [0, 0, 0, 0, 0, 0, 0];
let dailyRouteArray = [0, 0, 0, 0, 0, 0, 0];
let departmentSupportRouteArray = [0, 0, 0, 0, 0, 0, 0];
let pathVerificationArray = [0, 0, 0, 0, 0, 0, 0];
let signalIntegrityArray = [0, 0, 0, 0, 0, 0, 0]; 
let takedownArray = [0, 0, 0, 0, 0, 0, 0];
let dateArray = [];
let dates = [];
let newFileId;
let allArrays = [
  dates,
  dailyRouteArray,
  channelLaunchArray,
  takedownArray,
  pathVerificationArray,
  signalIntegrityArray,
  departmentSupportRouteArray,
];

function doGet() {
  const NEW_FILE = DriveApp.getFileById(FILE);
  //  Logger.log("Start");
  buildDateArray();
  let files = WEEKLY_REPORT_FOLDER.getFilesByName("" + dateArray[7] + " - " + dateArray[13] + " Weekly Routes Report");
  let HTMLString = "<style> h1, p {font-family: 'Helvetica', 'Arial'} </style>" + "<h1>" + dateArray[7] + " - " + dateArray[13] + " Weekly Routes Report";
  if (!files.hasNext()) {
    setNewFileId(NEW_FILE.makeCopy("" + dateArray[7] + " - " + dateArray[13] + " Weekly Routes Report").getId());
//    Logger.log(getNewFileId());
    setNewFileUrl(DriveApp.getFileById(getNewFileId()).getUrl());
//    Logger.log(getNewFileUrl());
    WEEKLY_REPORT_FOLDER.addFile(DriveApp.getFileById(getNewFileId()));
    HTMLString += " successfully created!</h1>";
    HTMLOutput = HtmlService.createHtmlOutput(HTMLString);
//    Logger.log("CREATED");
    //    return HTMLOutput;
  } else {
    HTMLString += " file already exists, no new file created.</h1>";
    HTMLOutput = HtmlService.createHtmlOutput(HTMLString);
//    Logger.log("FAILED");
    return HTMLOutput;
  }
  getDailyRoutes();
  getChannelLaunchRoutes();
  getDepartmentSupportRoutes();
  getPathVerificationRoutes();
  getSignalIntegrityRoutes();
  getTakedownRestoreRoutes();
//  Logger.log(getDailyRoutes() + "   $$$$$$$$$$");
//  Logger.log(getChannelLaunchRoutes() + "  %%%%%%%%");
//  Logger.log(getDepartmentSupportRoutes() + " !!!!!!!!!!");
//  Logger.log(getPathVerificationRoutes() + " ********");
//  Logger.log(getSignalIntegrityRoutes() + "  ^^^^^^^^^^");
//  Logger.log(getTakedownRestoreRoutes() + "  &&&&&&&&&");
  for (j = 0; j < 7; j++) {
    let toAddArray = [];
    let toAdd = [];
//    Logger.log(toAdd + " ************    " + i);
    for (i = 0; i < dates.length; i++) {
      toAdd = allArrays[j];
      toAddArray.push([toAdd[i]]);
    }
    transferToSpreadsheet(toAddArray, j);
  }
  SpreadsheetApp.openByUrl(getNewFileUrl()).getActiveSheet().getRange(3, 2, 1, 1).setValue("Week of " + dates[0] + " to " + dates[6]);

  Logger.log("........DONE.........");
  
  sendEmail();
  return HTMLOutput;
}

function buildDateArray() {
  year = date.getFullYear();
  today = parseInt(Utilities.formatDate(new Date(), "EST", "D"));
  for (i = 7; i > 0; i--) {
    dateArray.push(getDateFromDay(year, today, i));
    dates.push(getDateFromDayAltYear(year, today, i).toString());
  }
  let reg = new RegExp("(/)+", "gm");
  let len = dateArray.length;
  for (i = 0; i < len; i++) {
    temp = dateArray[i];
    dateArray.push(temp.replace(reg, "-"));
  }
//  for (i = 6; i >= 0; i--)
//  {
//    dateArray.push(getDateFromDayAlt(year, today, i));
//  }
  //  for (i = 6; i >= 0; i--)
  //  {
  //    dateArray.push(getDateFromDayAltYear(year, today, i));
  //  }
  for (i = 0; i < dateArray.length; i++) {
    Logger.log(dateArray[i]);
  }
}

function getDateFromDay(year, day, offset) {
  dateFromDay = new Date(year, 0, day);
  dateFromDay.setDate(dateFromDay.getDate() - offset);
  //    Logger.log(dateFromDay.toLocaleDateString('en-US', {
  //		month: "2-digit",
  //		day: "2-digit",
  //		year: "2-digit"
  //	}));
  return dateFromDay.toLocaleDateString("en-US", {
    month: "2-digit",
    day: "2-digit",
    year: "2-digit",
  });
}

function getDateFromDayAlt(year, day, offset) {
  dateFromDay = new Date(year, 0, day);
  dateFromDay.setDate(dateFromDay.getDate() - offset);
  return dateFromDay.toLocaleDateString("en-US", {
    month: "numeric",
    day: "numeric",
    year: "2-digit",
  });
}

function getDateFromDayAltMonth(year, day, offset) {
  dateFromDay = new Date(year, 0, day);
  dateFromDay.setDate(dateFromDay.getDate() - offset);
  return dateFromDay.toLocaleDateString("en-US", {
    month: "numeric",
    day: "2-digit",
    year: "2-digit",
  });
}

function getDateFromDayAltYear(year, day, offset) {
  dateFromDay = new Date(year, 0, day);
  dateFromDay.setDate(dateFromDay.getDate() - offset);
  return dateFromDay.toLocaleDateString("en-US", {
    month: "2-digit",
    day: "2-digit",
    year: "numeric",
  });
}

function getDateFromDayAltMonthAndYear(year, day, offset) {
  dateFromDay = new Date(year, 0, day);
  dateFromDay.setDate(dateFromDay.getDate() - offset);
  return dateFromDay.toLocaleDateString("en-US", {
    month: "numeric",
    day: "2-digit",
    year: "2-digit",
  });
}

function getMonthName(index) {
  // String array to convert numerical month to string
  const MONTHS = [
    "January", "February",
    "March", "April",
    "May", "June",
    "July", "August",
    "September", "October",
    "November", "December",
  ];
  return MONTHS[index];
}

function getDailyRoutes() {
  let folder;
  let checkFolder;
  let folderYear,
    folderMonth,
    dailyRoutesFolder,
    routes = 0;
  for (i = 0; i < dateArray.length; i++) {
    tempDate = dateArray[i];
    reg = new RegExp("(" + tempDate + ")" + "+");
    //    Logger.log(dateArray[i]);
    //    Logger.log(reg.toString());
    let month = getDateMonth(reg.toString());
    let checkMonth = "" + month + " - " + getMonthName(month - 1);
    //    Logger.log(checkMonth);
    // Starting from the daily log folder, find the year, then month
    folderYear = DriveApp.getFolderById(DAILY_LOGS_FOLDER).getFolders();
    while (folderYear.hasNext()) {
      folder = folderYear.next();
      // When found, get the Id and load to folderID
      if (folder.getName() === "" + date.getFullYear()) {
        checkFolder = folder.getId();
        folderId = DriveApp.getFolderById(folder.getId()).getFolders();
      }
    }
    folderMonth = DriveApp.getFolderById(checkFolder).getFolders();
    while (folderMonth.hasNext()) {
      folder = folderMonth.next();
      if (folder.getName() === checkMonth) {
        checkFolder = folder.getId();
        folderId = DriveApp.getFolderById(folder.getId());
      }
    }
    dailyRoutesFolder = DriveApp.getFolderById(checkFolder).getFiles();
    while (dailyRoutesFolder.hasNext()) {
      routes = getNumberOfRoutes(dailyRoutesFolder, i);
      insertIntoRouteArray(dailyRouteArray, i, routes);
    }
  }
  return dailyRouteArray;
}

function getChannelLaunchRoutes() {
  let chnLaunchFolder,
    routes = 0;
  //Iterate over the date array checking files in the channel launch folder
  for (i = 0; i < dateArray.length; i++) {
    //the folder containing the files to check
    chnLaunchFolder = DriveApp.getFolderById(CHANNEL_LAUNCH).getFiles();
    while (chnLaunchFolder.hasNext()) {
      //call the getNumberOfRoutes function to pull the data out of files
      routes = getNumberOfRoutes(chnLaunchFolder, i);
      insertIntoRouteArray(channelLaunchArray, i, routes);
    }
  }
  return channelLaunchArray;
}

function getDepartmentSupportRoutes() {
  let departmentFolder,
    routes = 0;
  //Iterate over the date array checking files in the channel launch folder
  for (i = 0; i < dateArray.length; i++) {
    //the folder containing the files to check
    departmentFolder = DriveApp.getFolderById(DEPARTMENT_FOLDER).getFiles();
    while (departmentFolder.hasNext()) {
      //call the getNumberOfRoutes function to pull the data out of files
      routes = getNumberOfRoutes(departmentFolder, i);
      insertIntoRouteArray(departmentSupportRouteArray, i, routes);
    }
  }
  return departmentSupportRouteArray;
}

function getPathVerificationRoutes() {
  let pathVerificationFolder,
    routes = 0;
  //Iterate over the date array checking files in the channel launch folder
  for (i = 0; i < dateArray.length; i++) {
    //the folder containing the files to check
    pathVerificationFolder = DriveApp.getFolderById(PATH_VERIFICATION).getFiles();
    while (pathVerificationFolder.hasNext()) {
      //call the getNumberOfRoutes function to pull the data out of files
      routes = getNumberOfRoutes(pathVerificationFolder, i);
      insertIntoRouteArray(pathVerificationArray, i, routes);
    }
  }
  return pathVerificationArray;
}

function getSignalIntegrityRoutes() {
  let signalIntegrityFolder,
    routes = 0;
  //Iterate over the date array checking files in the channel launch folder
  for (i = 0; i < dateArray.length; i++) {
    //the folder containing the files to check
    signalIntegrityFolder = DriveApp.getFolderById(SIGNAL_INTEGRITY).getFiles();
    while (signalIntegrityFolder.hasNext()) {
      //call the getNumberOfRoutes function to pull the data out of files
      routes = getNumberOfRoutes(signalIntegrityFolder, i);
      insertIntoRouteArray(signalIntegrityArray, i, routes);
    }
  }
  return signalIntegrityArray;
}

function getTakedownRestoreRoutes() {
  let takedownRestoreFolder,
    routes = 0;
  //Iterate over the date array checking files in the channel launch folder
  for (i = 0; i < dateArray.length; i++) {
    //the folder containing the files to check
    takedownRestoreFolder = DriveApp.getFolderById(TAKEDOWN_RESTORE).getFiles();
    while (takedownRestoreFolder.hasNext()) {
      //call the getNumberOfRoutes function to pull the data out of files
      routes = parseInt(getNumberOfRoutes(takedownRestoreFolder, i));
      insertIntoRouteArray(takedownArray, i, routes);
    }
  }
  return takedownArray;
}

function getDateDay(dateStr) {
  return dateStr.slice(5, 7);
}

function getDateMonth(dateStr) {
  return dateStr.slice(2, 4);
}

function getDateYear(dateStr) {
  return "20" + dateStr.slice(8, 10);
}

function getNumberOfRoutes(folder, i) {
  let temp,
    re,
    file,
    result,
    routeCount = 0;
  temp = dateArray[i];
  //Regular expression of a date to check files for
  re = new RegExp("(" + temp + ")" + "+");
  //  Logger.log(re.toString());
  file = folder.next();
  //The regex test
  result = re.test(file);
  //  if (i > 6 && i < 14)
  //  {
  //    Logger.log(getDateMonth(re.toString()));
  //    Logger.log(getDateDay(re.toString()));
  //    Logger.log(getDateYear(re.toString()));
  //  }
  //Open the sheet and get the value out of this specific cell
  if (result) {
    routeCount = SpreadsheetApp.open(file).getRange("D2").getValue();
  }
  return routeCount;
}

function insertIntoRouteArray(array, i, routes) {
  switch (i) {
    case 0:
    case 7:
    case 14:
    case 21:
      array[0] += routes;
      break;
    case 1:
    case 8:
    case 15:
    case 22:
      array[1] += routes;
      break;
    case 2:
    case 9:
    case 16:
    case 23:
      array[2] += routes;
      break;
    case 3:
    case 10:
    case 17:
    case 24:
      array[3] += routes;
      break;
    case 4:
    case 11:
    case 18:
    case 25:
      array[4] += routes;
      break;
    case 5:
    case 12:
    case 19:
    case 26:
      array[5] += routes;
      break;
    case 6:
    case 13:
    case 20:
    case 27:
      array[6] += routes;
      break;
  }
}

function transferToSpreadsheet(data, j) {
  SpreadsheetApp.openById(getNewFileId()).getActiveSheet().getRange(5, 2 + j, data.length, 1).setValues(data);
}

function setNewFileId(id) {
  newFileId = id;
}

function getNewFileId() {
  return newFileId;
}

function setNewFileUrl(url) {
  newFileUrl = url;
}

function getNewFileUrl() {
  return newFileUrl;
}

function sendEmail() {
  // Send an email with a link to the file on Google Drive.
  let fileToSend = getNewFileUrl();
  let start = {
    date: dateArray[0]
  };
  let end = {
    date: dateArray[6]
  };
  let file = {
    url: fileToSend
  };
  let templ = HtmlService
      .createTemplateFromFile('Email');
  templ.start = start;
  templ.end = end;
  templ.file = file;
  let message = templ.evaluate().getContent();
  GmailApp.sendEmail('WHO_TO_EMAIL_TO', 'Weekly Routes Report', 'Please follow the link for this week\'s route report ' + fileToSend, { //  Cheyenne-TOCTeamLeaders@dish.com
//    htmlBody: 'Please follow the link for this week\'s route report ' + '<a href=\"' + fileToSend + '">Weekly Route Report</a>',
    htmlBody: message,
    name: 'Route Report'
  });
}
