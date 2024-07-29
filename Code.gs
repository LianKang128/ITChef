function doGet(e) {
  // Fetch and populate data when the web app is accessed
  readFirebaseData();

  // Return the main HTML page
  return HtmlService.createTemplateFromFile('index').evaluate();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getPage(page) {
  return HtmlService.createHtmlOutputFromFile(page).getContent();
}

function readFirebaseData() {
  const customerUrl = 'https://g-workspace-hackathon-default-rtdb.asia-southeast1.firebasedatabase.app/Customers.json';
  const productUrl = 'https://g-workspace-hackathon-default-rtdb.asia-southeast1.firebasedatabase.app/Products.json';

  try {
    const customerResponse = UrlFetchApp.fetch(customerUrl);
    const customerData = JSON.parse(customerResponse.getContentText());
    Logger.log('Fetched Customer Data: ' + JSON.stringify(customerData));

    const productResponse = UrlFetchApp.fetch(productUrl);
    const productData = JSON.parse(productResponse.getContentText());
    Logger.log('Fetched Product Data: ' + JSON.stringify(productData));

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const customerSheet = ss.getSheetByName('Customers');
    const productSheet = ss.getSheetByName('Products');
    const pastPurchaseSheet = ss.getSheetByName('Past Purchases');

    // Clear existing content below headers
    customerSheet.getRange(2, 1, customerSheet.getMaxRows() - 1, customerSheet.getMaxColumns()).clearContent();
    productSheet.getRange(2, 1, productSheet.getMaxRows() - 1, productSheet.getMaxColumns()).clearContent();
    pastPurchaseSheet.getRange(2, 1, pastPurchaseSheet.getMaxRows() - 1, pastPurchaseSheet.getMaxColumns()).clearContent();

    // Populate customer data
    let customerRow = 2; // Start from the second row
    for (let key in customerData) {
      Logger.log('Processing Customer: ' + key);
      const row = [
        key,
        customerData[key].Name,
        customerData[key].Email,
        customerData[key].DOB,
        customerData[key].Gender,
        customerData[key].Nationality
      ];
      Logger.log('Appending Row: ' + JSON.stringify(row));
      customerSheet.getRange(customerRow++, 1, 1, row.length).setValues([row]);
    }

    // Populate product data
    let productRow = 2; // Start from the second row
    for (let key in productData) {
      const row = [
        key, // Assuming key is Product ID
        productData[key].Name,
        productData[key].Price
      ];
      Logger.log('Appending Row: ' + JSON.stringify(row));
      productSheet.getRange(productRow++, 1, 1, row.length).setValues([row]);
    }

    // Note: Add code to fetch and populate past purchase data if available
  } catch (error) {
    Logger.log('Error fetching data: ' + error);
  }
}

function sendEmails(criteria) {
  Logger.log("Received criteria: " + JSON.stringify(criteria));

  // Check if the criteria object has the expected properties
  if (!criteria || typeof criteria !== 'object') {
    Logger.log("Invalid criteria object");
    throw new Error("Invalid criteria object");
  }

  if (!criteria.subject || !criteria.body) {
    Logger.log("Subject or body is missing");
    throw new Error("Email subject and body cannot be empty.");
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  let selectedCustomers = data;

  if (criteria.sendAllCustomers) {
    Logger.log("Send to all customers selected");
    // No filtering needed, send to all customers
  } else {
    if (criteria.birthdayMonth) {
      Logger.log("Filtering by birthday month");
      const currentMonth = new Date().getMonth() + 1;
      selectedCustomers = selectedCustomers.filter(row => {
        let dob = row[3]; // Assuming DOB is in the 4th column in format YYYYMMDD
        Logger.log("DOB value: " + dob + " (Type: " + typeof dob + ")");
        
        if (typeof dob === 'number') {
          dob = dob.toString(); // Convert numeric DOB to string
        }

        if (typeof dob === 'string' && dob.length >= 6) {
          return parseInt(dob.substring(4, 6)) === currentMonth;
        } else {
          Logger.log("Invalid DOB format: " + dob);
          return false;
        }
      });
    }

    if (criteria.ageFrom || criteria.ageTo) {
      Logger.log("Filtering by age range");
      const ageFrom = criteria.ageFrom ? parseInt(criteria.ageFrom) : 0;
      const ageTo = criteria.ageTo ? parseInt(criteria.ageTo) : 100;
      selectedCustomers = selectedCustomers.filter(row => {
        let dob = row[3]; // Assuming DOB is in the 4th column in format YYYYMMDD
        Logger.log("DOB value: " + dob + " (Type: " + typeof dob + ")");
        
        if (typeof dob === 'number') {
          dob = dob.toString(); // Convert numeric DOB to string
        }

        if (typeof dob === 'string' && dob.length >= 8) {
          const birthYear = parseInt(dob.substring(0, 4));
          const age = new Date().getFullYear() - birthYear;
          return age >= ageFrom && age <= ageTo;
        } else {
          Logger.log("Invalid DOB format: " + dob);
          return false;
        }
      });
    }

    if (criteria.gender) {
      Logger.log("Filtering by gender");
      selectedCustomers = selectedCustomers.filter(row => row[4] === criteria.gender);
    }

    if (criteria.nationality.length > 0) {
      Logger.log("Filtering by nationality");
      selectedCustomers = selectedCustomers.filter(row => criteria.nationality.includes(row[5]));
    }
  }

  Logger.log("Number of selected customers: " + selectedCustomers.length);

  const emailSubject = criteria.subject;
  const emailBody = criteria.body;

  const attachments = criteria.files.map(file => {
    return {
      fileName: file.name,
      mimeType: file.type,
      content: Utilities.base64Decode(file.content)
    };
  });

  selectedCustomers.forEach(customer => {
    const email = customer[2]; // Assuming Email is in the 3rd column
    MailApp.sendEmail({
      to: email,
      subject: emailSubject,
      htmlBody: emailBody,
      attachments: attachments
    });
  });

  Logger.log("Emails sent successfully");

  // Return the count of emails sent
  return selectedCustomers.length;
}

// function analyzeCustomerData() {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
//   const data = sheet.getDataRange().getValues();
//   const headers = data.shift();

//   // Perform analysis (e.g., customer segmentation, purchase trends)
//   // Placeholder for analysis logic

//   // Write analysis results back to the sheet
//   sheet.appendRow(['Analysis Results']);
//   // Example: Add more rows as needed for your analysis results
// }

// function recommendProducts() {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Past Purchases');
//   const data = sheet.getDataRange().getValues();
//   const headers = data.shift();

//   // Placeholder for recommendation logic

//   // Example: Add recommendations back to the sheet
//   sheet.appendRow(['Recommendations']);
//   // Example: Add more rows as needed for your recommendations
// }

// function onFormSubmit(e) {
//   const responses = e.values;
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

//   // Process form responses (e.g., customer feedback)
//   // Placeholder for processing logic

//   // Write feedback to the sheet
//   sheet.appendRow(['Customer Feedback']);
//   // Example: Add more rows as needed for feedback processing
// }