const userEmail = Session.getActiveUser().getEmail();
const ss = SpreadsheetApp.getActive();
let as = ss.getActiveSheet();
const sheetName = as.getSheetName();
const startRow = 6;
const startCol = 1;
const lastRow = as.getLastRow();
const lastCol = as.getLastColumn();

// Template Sheet Data
const templateSheetName = "Template";
let templateSheet = ss.getSheetByName(templateSheetName);
const templateVal = templateSheet.getRange("A2").getDisplayValue();

// Active Sheet Data
const asVal = as.getRange("A2").getDisplayValue();
const calID = as.getRange("A3").getDisplayValue();
const exclude = as
  .getRange("B3")
  .getDisplayValue()
  .split(",")
  .map((item) => item.trim());

function getCalendarEvents() {
  const { firstDate, lastDate } = getDefaultDate();

  const calendar = CalendarApp.getCalendarById(userEmail);
  const calendarEvents = calendar.getEvents(firstDate, lastDate);

  if (calendarEvents === undefined) {
    throw {
      title: "Error while getting calendar events!",
      prompt: `There is no activity on "${userEmail}" calendar!\n\nTry to add activity first`,
    };
  }

  return calendarEvents;
}

function getActivitiesData() {
  const calendarEvents = getCalendarEvents();

  // Create a Map to store the data that sum the duration if activity exists multiple time within same week
  let weeklyActivitiesMap = new Map();

  // Loop through calendarEvents and get the details
  calendarEvents.forEach((event) => {
    // Get default values from calendar
    const myStatus = event.getMyStatus().toString(); // If you decline an event invitation in the calendar, it will be excluded
    const title = event.getTitle();
    if (!title || title.trim() === "") {
      return;
    }
    const startTime = event.getStartTime();
    const endTime = event.getEndTime();

    if (
      exclude.some(
        (excludedTitle) =>
          title.trim().toLowerCase() === excludedTitle.toLowerCase()
      ) ||
      !["YES", "OWNER"].includes(myStatus)
    ) {
      return;
    }

    const week = getISOWeekNumber(startTime);
    const dateStr = `${startTime.getDate()}/${
      startTime.getMonth() + 1
    }/${startTime.getFullYear()}`; // output this date format "d/mm/yyyy"
    const duration = (endTime.getTime() - startTime.getTime()) / 1000 / 60 / 60;

    // This step is to make sure if the same activity exists within the same week, it combines the duration
    const key = `${title}-${week}`;

    // Check if the key already exists
    if (weeklyActivitiesMap.has(key)) {
      // If it does, add the duration to get the total duration
      let value = weeklyActivitiesMap.get(key);
      value[3] += duration;
      weeklyActivitiesMap.set(key, value);
    } else {
      // If it doesn't, then we add the activity to the map object
      weeklyActivitiesMap.set(key, [dateStr, week, title, duration]);
    }
  });

  // Convert to list to be used
  const activitiesData = [...weeklyActivitiesMap.values()];

  return activitiesData;
}

function setDataToSheet() {
  const activitiesData = getActivitiesData();
  let numberOfActivities;
  let statusMessage;

  // Update current active sheet
  as = ss.getActiveSheet();

  // Set format for some ranges
  as.getRange(`A${startRow}:A`).setNumberFormat("d/mm/yyyy");
  as.getRange(`B${startRow}:B`)
    .setNumberFormat("#")
    .setHorizontalAlignment("center");
  as.getRange(`D${startRow}:D`)
    .setNumberFormat("#,##0.00")
    .setHorizontalAlignment("center");

  // If there is no activity in the sheet yet, write the first list of activitiesData, else append new activitiesData
  if (lastRow === startRow - 1) {
    as.getRange(
      lastRow + 1,
      1,
      activitiesData.length,
      activitiesData[0].length
    ).setValues(activitiesData);
    numberOfActivities = activitiesData.length;
  } else {
    // Get activitiesData to compared
    const newActivitiesData = activitiesData.map((activity) =>
      [activity[1], activity[2]].join(",")
    );
    const oldActivitiesData = as
      .getRange(startRow, 2, lastRow - (startRow - 1), 2)
      .getDisplayValues()
      .map((activity) => activity.join(","));

    // Find missed and edited activitiesData
    const missedActivities = [];

    newActivitiesData.forEach((key, index) => {
      const activitiesIndex = activitiesData[index];
      if (!oldActivitiesData.includes(key)) {
        missedActivities.push(activitiesIndex);
      } else {
        // Update edited activitiesData value in the sheet
        const rowIndex = oldActivitiesData.indexOf(key); // get the oldActivitiesData index that match the key
        as.getRange(rowIndex + startRow, 5).setValue(activitiesIndex[3]); // Add actual duration
      }
    });

    // If there is missed activitiesData, add it to the sheet
    if (missedActivities.length > 0) {
      as.getRange(
        lastRow + 1,
        1,
        missedActivities.length,
        missedActivities[0].length
      ).setValues(missedActivities);
      numberOfActivities = missedActivities.length;
    }
  }

  // Toest a message depending on the updated data
  if (numberOfActivities > 0) {
    statusMessage = `Succesfully add ${numberOfActivities} activitiesData to "${sheetName}" sheet`;
  } else {
    statusMessage = `Succesfully update activitiesData to "${sheetName}" sheet`;
  }

  // Reorder the activitiesData based on date ascending and set the border style
  as.getRange(`A${startRow}:E`)
    .setVerticalAlignment("top")
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      "#999999",
      SpreadsheetApp.BorderStyle.SOLID
    )
    .sort({ column: 2, ascending: false });

  return statusMessage;
}

function inspectSheet() {
  let errTitle;
  let errMessage;

  // Catch error if the active sheet is not a valid workload sheet
  if (asVal !== templateVal) {
    errTitle = "Sheet structure does not match!";
    errMessage =
      "Please do not edit the template sheet structure or Make sure to update your activitiesData from a valid workload sheet";
    throw { title: errTitle, prompt: errMessage };
  }

  // Catch the error if sheet are the master template
  if (sheetName === templateSheetName) {
    errTitle = `${sheetName} is a template sheet!`;
    errMessage =
      "Please do not use this sheet! If you does not have workload sheet monitoring yet, create one using the menu button";
    throw { title: errTitle, prompt: errMessage };
  }

  // Catch the error if calID doesn't match userEmail
  if (calID === "") {
    errTitle = `Your Calendar ID is still empty!`;
    errMessage = `This is your default calendar ID (${userEmail}`;
    throw { title: errTitle, prompt: errMessage };
  } else if (calID !== userEmail) {
    errTitle = `Your email does not match!`;
    errMessage = `Your current account (${userEmail}) does not match with (${sheetName}) sheet "Calendar ID"!\n\nMake sure to update your activitiesData from your own workload sheet or copy (${userEmail}) to the "Calendar ID"`;
    throw { title: errTitle, prompt: errMessage };
  }
}

function updateUserWorkloadSheet() {
  // Validate the current active sheet
  inspectSheet();

  const statusMessage = setDataToSheet();
  ss.toast(statusMessage, "Status", 5);
}

// Copy the template sheet and name it as inputted by the user
function createNewWorkloadSheet(strInput) {
  templateSheet.copyTo(ss).setName(strInput);
  const copiedSheet = ss.getSheetByName(strInput);
  copiedSheet.getRange("A3").setValue(userEmail);
  ss.setActiveSheet(copiedSheet);
  SpreadsheetApp.flush();
}
