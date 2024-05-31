/**
 * How to use this script:
 * 1. Open Google Sheets:
 *    - Go to your Google Sheets document.
 *
 * 2. Open Script Editor:
 *    - Click on `Extensions` in the menu.
 *    - Select `Apps Script`.
 *
 * 3. Create a New Script:
 *    - In the Apps Script editor, delete any code in the script editor and paste this script.
 *    - Save the project with a name, for example, `Password Generator`.
 *
 * 4. Authorize the Script:
 *    - The first time you run the script, you will need to authorize it to access your Google Sheets.
 *    - Click on the play button (triangle icon) to run the `fillColumnWithPasswords` function.
 *    - Follow the authorization prompts.
 *
 * 5. Run the Script:
 *    - After authorization, you can run the script by selecting "Generate Passwords" from the "Password Generator" menu in your Google Sheets.
 *
 * What the Script Does:
 * - Generate Strong Passwords: The script generates strong passwords that are at least the specified number of characters long, including uppercase and lowercase letters, numbers, and special characters.
 * - Fill Specified Range in Google Sheets: It prompts the user for the column letter, the range of rows, and the password length where the passwords should be added, then fills the specified cells with the generated passwords.
 *
 * Created by Jonas Lund 2024
 */

/**
 * Adds a custom menu to the Google Sheets UI when the document is opened.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Password Generator')
    .addItem('Generate Passwords', 'fillColumnWithPasswords')
    .addToUi();
}

/**
 * Prompts the user for the column letter, the range of rows, and the password length, then fills the specified cells with generated passwords.
 * Created by Jonas Lund 2024
 */
function fillColumnWithPasswords() {
  var ui = SpreadsheetApp.getUi();
  
  // Prompt the user for inputs
  var columnLetter = ui.prompt('Column Selection', 'Enter the column letter to add passwords:', ui.ButtonSet.OK).getResponseText().toUpperCase();
  var startRow = parseInt(ui.prompt('Start Row', 'Enter the start row number:', ui.ButtonSet.OK).getResponseText());
  var endRow = parseInt(ui.prompt('End Row', 'Enter the end row number:', ui.ButtonSet.OK).getResponseText());
  var length = parseInt(ui.prompt('Password Length', 'Enter the desired password length (minimum 12 characters):', ui.ButtonSet.OK).getResponseText());

  // Validate inputs
  if (isNaN(startRow) || isNaN(endRow) || isNaN(length) || startRow <= 0 || endRow < startRow || length < 12) {
    ui.alert('Invalid input. Ensure row numbers and password length are positive integers, end row is greater than or equal to start row, and password length is at least 12.');
    return;
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rowCount = endRow - startRow + 1;
  var passwords = generatePasswordArray(rowCount, length);
  var columnIndex = columnLetter.charCodeAt(0) - 'A'.charCodeAt(0) + 1;
  var columnRange = sheet.getRange(startRow, columnIndex, passwords.length, 1);
  columnRange.setValues(passwords);
}

/**
 * Generates an array of passwords with the specified number of rows and length.
 * @param {number} rows The number of rows.
 * @param {number} length The desired length of each password.
 * @return {Array<Array<string>>} The array of passwords.
 */
function generatePasswordArray(rows, length) {
  var passwordArray = [];
  for (var i = 0; i < rows; i++) {
    passwordArray.push([generatePassword(length)]);
  }
  return passwordArray;
}

/**
 * Generates a random password with at least the specified number of characters, including lowercase, uppercase, numbers, and special characters.
 * @param {number} length The desired length of the password.
 * @return {string} The generated password.
 */
function generatePassword(length) {
  var lowercase = "abcdefghijklmnopqrstuvwxyz";
  var uppercase = lowercase.toUpperCase();
  var numbers = "0123456789";
  var special = "!#4%&?$*";
  var allChars = lowercase + uppercase + numbers + special;

  var password = [];
  password.push(randomChar(numbers));  // Ensure at least one number
  password.push(randomChar(special));  // Ensure at least one special character

  while (password.length < length) {
    password.push(randomChar(allChars));
  }

  return shuffleString(password.join(''));
}

/**
 * Returns a random character from the given string.
 * @param {string} str The string to pick a random character from.
 * @return {string} The random character.
 */
function randomChar(str) {
  return str.charAt(Math.floor(Math.random() * str.length));
}

/**
 * Shuffles the characters in the given string.
 * @param {string} str The string to shuffle.
 * @return {string} The shuffled string.
 */
function shuffleString(str) {
  return str.split('').sort(function() { return 0.5 - Math.random(); }).join('');
}
