/*******************************************************************************************
 *                                                                                         *
 *               Update Completed Credits with Fall/Spring CR-CA Data 	                   *
 *                                                                                         *
 *       Alvaro Gomez                                                                      *
 *       Academic Technology Coach                                                         *
 *       Northside Independent School District                                             *
 *       alvaro.gomez@nisd.net                                                             *
 *       +1-210-363-1577                                                                   *
 *                                                                                         *
 * Purpose: To automate the task of copying over rows of data from one Google Sheet        *
 *          into another. Also to send out an email to the student's counselor             *
 *          letting them know their student completed a particular course.                 *
 *                                                                                         *
 * The functions below do three things:                                                    *
 *      1. onOpen creates a User Interface in the '24-25 OC AT-RISK CR/CA DATABASE'        *
 *         Google Spreadsheet. The user interface has one option, which is to run the      *
 *         updateCompletedCredits function.                                                *
 *                                                                                         *
 *      2. showConfirmationDialog is a simple confirmation to the user who selects to      *
 *         update the Completed Credits sheet.                                             *
 *                                                                                         *
 *      3. updateCompletedCredits looks for rows in the "Fall/Spring CR-CA Data" sheet     *
 *         that contain checks (are TRUE) in column A. If a check is there, then it will   *
 *         look to see if that row doesn't already exist in "Completed Credits". If it     *
 *         doesn't exist then it will insert the row into "Completed Credits". At the end  *
 *         of this function is calls the sendCounselorNotification function.               *
 *                                                                                         *
 *      4. sendCounselorNotification sends an email to the student's counselor to let them *
 *         know that the student completed a course.                                       *
 *                                                                                         *
 * File formats: The function reads the data from a sheet within a Google Spreadsheet      *
 *               and writes the specific rows to another sheet within that same            *
 *               spreadsheet.                                                              *
 *                                                                                         *
 * Revision history: 8/27/24-Alvaro Gomez-Fixed data imports                               *
 *                                                                                         *
*******************************************************************************************/

// Creates the menu with the options to import new students from the Fall/Spring CR-CA Data sheet or to watch the instructions video
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('UPDATE the Completed Credits Sheet')
    .addItem('Import new students from Fall/Spring CR-CA Data', 'showConfirmationDialog')
    .addItem('Watch instructions video', 'openInstructionsVideo')
    .addToUi();
}

// Shows a confirmation dialog to the user before updating the Completed Credits sheet
function showConfirmationDialog() {
  let ui = SpreadsheetApp.getUi();
  let response = ui.alert(
    'UPDATE the Completed Credits sheet',
    'Click yes to begin inserting new rows from the Fall/Spring CR-CA Data sheet into the Completed Credits sheet.',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    updateCompletedCredits();
    ui.alert('Completed Credits updated successfully.');
  }
}

function openInstructionsVideo() {
  let videoUrl = 'https://youtu.be/L9qvNBAhtBw';
  let htmlOutput = HtmlService.createHtmlOutput('<script>window.open("' + videoUrl + '");google.script.host.close();</script>');
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Tutorial Video');
}

function updateCompletedCredits() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let autumnSpringDatasheet = spreadsheet.getSheetByName("Fall/Spring CR-CA Data");
  let completedCreditsSheet = spreadsheet.getSheetByName("Completed Credits");
  let autumnSpringDataRange = autumnSpringDatasheet.getRange("A2:O" + autumnSpringDatasheet.getLastRow());
  let existingDataRange = completedCreditsSheet.getRange("A2:M" + completedCreditsSheet.getLastRow());
  let existingData = existingDataRange.getValues();
  let newData = [];
  let counselorNotification = [];
  let autumnData = autumnSpringDataRange.getValues();

  for (let i = 0; i < autumnData.length; i++) {
    let row = autumnData[i];
    let isChecked = row[0]; // Checkbox in column A
    let studentName = row[5].toString().toLowerCase(); // Student Name in column F
    let studentID = row[6]; // Student ID in column G
    let courseName = row[7].toString().toLowerCase(); // Course Name in column H
    let courseNo = row[8]; // Course No in column I
    let courseDateStart = row[9]; // Course Date Start in column J
    let courseEndStart = row[10]; // CourseEndDate Start in column K
    let courseGrade = row[11]; // CourseGrade in column L
    let timeoncourse = row[13]; // timeoncourse in column N
    let url = row[14]; // LOC link in column O

    if (!isChecked) {
      continue; // Skip student if column A is not checked
    }

    let isDuplicate = false;

    for (let j = 0; j < existingData.length; j++) {
      let existingRow = existingData[j];
      let existingStudentID = existingRow[2];
      let existingCourseName = existingRow[3].toString().toLowerCase();

      if (
        existingStudentID === studentID &&
        existingCourseName === courseName
      ) {
        isDuplicate = true;
        break;
      }
    }

    if (isDuplicate) {
      continue; // Skip if student ID and course are duplicates
    }

    let newRow = [
      row[5].toString().toUpperCase(), // Student Name in Column F
      studentID, // Student ID in Column G
      courseName.toUpperCase(), // Course Name in Column G
      courseNo,
      //parseInt(courseNo,10), // Course ID in Column I
      courseDateStart, // Course Date Start in Column J
      row[10], // Course Date Credit Earned from Column K
      row[11],
      //row[11] !== "" && !isNaN(parseInt(row[10])) ? parseInt(row[10]) : "", // Course Grade Average in Column L
      row[12], // Teacher of Record in Column M
      row[13], // Hours on course if CA-MT completion from Column N
      row[14], // URL of the completion letter from Column O
      "", // LS
      "" // NOTES
    ];

    sendCounselorNotification(row[5], studentID, courseName.toUpperCase())
    newData.push(newRow);
  }

  newData.sort(function(a, b) {
    let studentA = String(a[1]).toLowerCase();
    let studentB = String(b[1]).toLowerCase();
    return studentA.localeCompare(studentB);
  });

  let insertIndex = 2;

  for (let m = 0; m < newData.length; m++) {
    let studentID = String(newData[m][1]).toLowerCase();
    let nextStudentID = (m + 1 < newData.length) ? String(newData[m + 1][1]).toLowerCase() : "";

    if (studentID.localeCompare(nextStudentID) < 0) {
      insertIndex++;
    }

    completedCreditsSheet.insertRowBefore(insertIndex);
    completedCreditsSheet.getRange(insertIndex, 2, 1, newData[m].length).setValues([newData[m]]).setBackground(null).setBorder(true, true, true, true, true, true);
    insertIndex++;
  }

  completedCreditsSheet.getRange("A2:M" + completedCreditsSheet.getLastRow()).sort({ column: 2, ascending: true });

  // Renumber rows in column A starting from 1
  let lastRow = completedCreditsSheet.getLastRow();
  let rowNumbers = completedCreditsSheet.getRange("A2:A" + lastRow).getValues();
  let updatedRowNumbers = rowNumbers.map(function (value, index) {
    return [index + 1];
  });
  completedCreditsSheet.getRange("A2:A" + lastRow).setValues(updatedRowNumbers);

}

function sendCounselorNotification(student, id, course) {
  // Map of counselor names to their corresponding emails
  let counselorEmails = {
    '(A-C) Appleby': 'janelle.appleby@nisd.net',
    '(D-Ha) Hewgley': 'shanna.hewgley@nisd.net',
    '(He-Mi) Ramos': 'elizabeth.ramos@nisd.net',
    '(Mo-R) Clarke': 'darrell.clarke@nisd.net',
    '(S-Z) Pearson': 'samantha.pearson@nisd.net',
  };

/****************************************************
 * Use the map below for testing purposes. This is  *
 * a map of the counselor names going to Alvaro's   *
 * email.                                           *
 *                                                  *
 * When testing, uncomment the map below and        *
 * comment out the map above.                       *
****************************************************/
  // let counselorEmails = {
  //   '(A-C) Appleby': 'alvaro.gomez@nisd.net',
  //   '(D-Ha) Hewgley': 'alvaro.gomez@nisd.net',
  //   '(He-Mi) Ramos': 'alvaro.gomez@nisd.net',
  //   '(Mo-R) Clarke': 'alvaro.gomez@nisd.net',
  //   '(S-Z) Pearson': 'alvaro.gomez@nisd.net',
  // };

  // Get the first letter of the student's last name
  let lastNameFirstTwoLetters = student.split(",")[0].trim().substring(0,2).toUpperCase();

  // Determine the counselor based on the first two letters of the student's last name
  let counselor;
  if (lastNameFirstTwoLetters >= 'AA' && lastNameFirstTwoLetters <= 'CZ') {
    counselor = '(A-C) Appleby';
  } else if (lastNameFirstTwoLetters >= 'DA' && lastNameFirstTwoLetters <= 'HA') {
    counselor = '(D-Ha) Hewgley';
  } else if (lastNameFirstTwoLetters >= 'HE' && lastNameFirstTwoLetters <= 'MI') {
    counselor = '(He-Mi) Ramos';
  } else if (lastNameFirstTwoLetters >= 'MO' && lastNameFirstTwoLetters <= 'RZ') {
    counselor = '(Mo-R) Clarke';
  } else if (lastNameFirstTwoLetters >= 'SA' && lastNameFirstTwoLetters <= 'ZZ') {
    counselor = '(S-Z) Pearson';
  } else {
    counselor = '(Head Counselor) Matta'; // default counselor
  }

  var recipient = counselorEmails[counselor]; // get the counselor's email;
  var subject = "Student Completed CR/CA";
  var senderEmail = "angela.guajardo@nisd.net, katherine.popp@nisd.net";
  var body = `Dear Counselor,\n\nWe are happy to report ${student} (${id}), has completed: ${course}\n\nWhat should they work on next or are they all done?\n\nThank you,\nMs. Guajardo and Mrs. Popp`;
  
  MailApp.sendEmail({
      to: recipient,
      subject: subject,
      body: body,
      from: senderEmail
  });

}