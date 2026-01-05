/*
 * @module Conan Rental Management Automation
 * @author Michael Conan
 * @version 2.0
 * @date 2022-10-15
 * @description Scripts to pull utility bills from email, post recurring transactions and send update for main
 *              home and Duplex rental property. Additionally handles expense submissions from Google form.
 *              Dependent on associated Google sheets (for each property) and utility configuration file.
 */

// ~~ GLOBALS ~~
// try/catch for limited auth functions
try {
  // Testing flag for unit tests
  var TEST = false;
  
  // Globals to configure
  var PROPS = PropertiesService.getScriptProperties();
  var MAIN_USER = Session.getEffectiveUser().getEmail();
  var RECIPIENTS = PROPS.getProperty('recipients');

  // Utilities key config
  var UTILITIES = JSON.parse(DriveApp.getFileById('1j9y-WgXFvDztALYFMZr68WnLTapuhXyE').getBlob().getDataAsString());
  var UTILITY_FOLD = DriveApp.getFolderById('1GqO-9jl-vX4JSD00jO3OmbVIjyhB56ss');

  // CRM Spreadsheet
  var CRM = SpreadsheetApp.getActive();
  var LEDGER = CRM.getSheetByName('Transaction Detail');

  // Dates / Times
  var TZ = CRM.getSpreadsheetTimeZone();
  var MONTH = Utilities.formatDate(new Date, TZ, "yyyy MMMM");
  var RENTDATE = new Date(new Date().getTime() + 4*24*60*60000);

  // Rommate Utilities Spreadsheet (Deprecated)
  var SS = SpreadsheetApp.openById('1JAT7ZaptjyHrEDLAt9RhWfSEJ60RvX0_vPCU-8XFK-I');
  var SHEET = SS.getSheetByName('Sheet1');

  // Main Spreadsheet
  var SMAIN = SpreadsheetApp.openById('1JwiQPOouopVfoE4wB0_lnGbq-8u-xBTDTNAEhR-AMJw');
  var TRANS = SMAIN.getSheetByName('Transactions');
    
  // Utility mail labels
  var BILLS = GmailApp.getUserLabelByName('Home/Bills').getThreads(0,100);
  var RECORDS = GmailApp.getUserLabelByName('Home/Duplex Records').getThreads(0,100);
  var REVIEW = GmailApp.getUserLabelByName('Script/Unmatched');
  var ERROR = GmailApp.getUserLabelByName('Script/Error');
}
catch (e) {
  Logger.log('Some globals cant be loaded from simple trigger... '+ e);
}

/**
 * Main orchestration function for montly property management activities
 */
function manageProperty() {
  
  // authorize
  if (!auth_()) {
      return;
  }
  
  // Run property mangement functions
  chargeUtilities(); // Old function for roommate utility sharing
  //getBills(); // Separate trigger due to runtime
  getRent();
  getMortgage();
  closeBooks();
  (TEST) ? sendUpdates(MAIN_USER) : sendUpdates(RECIPIENTS);
}

/**
 * Function to iterate bill accounts from config, identify emails from labels (filtered by Gmail),
 * parse amounts and post entries to spreadsheets
 */
function getBills(scope) {
  
  // authorize
  if (!auth_()) {
      return;
  }
  
  // Add transactions to ledger for any emails not previously logged
  let tMain = TRANS.getDataRange().getValues();
  for (let u in UTILITIES) {
    if (scope && !scope.includes(u)) {
      continue
    }
    var mails = RECORDS.filter(m => m.getFirstMessageSubject().includes(UTILITIES[u].subject) && m.getMessages()[0].getFrom().indexOf(UTILITIES[u].from)+1);
    var billMails = BILLS.filter(m => m.getFirstMessageSubject().includes(UTILITIES[u].subject) && m.getMessages()[0].getFrom().indexOf(UTILITIES[u].from)+1);
    mails = mails.concat(billMails);
    Logger.log(mails.length + " bill mails for " + u);
    for (let m of mails) {
      let match = true;
      try {
        var msg = m.getMessages()[0];

        // Supports PDF file attachment, HTML body or plain body based on configurations
        var body = (UTILITIES[u].attachment) ? getAttachmentBody_(msg) : (UTILITIES[u].html) ? msg.getBody() : msg.getPlainBody();
        
        // If payment confirmation, use message date, otherwise parse due date from bill
        var mDate = (UTILITIES[u].payment) ? msg.getDate() : body.match(/Due Date.*(\d{2}\/\d{2}\/\d{2})/)[1];
        
        // Check account number before posting
        if (body.indexOf(UTILITIES[u].account)+1) {
          // Don't add bills to CRM sheet
          if (UTILITIES[u].label == 'bills') {
            continue;
          }
          // Parse utility amount and attempt to post (will ignore duplicates)
          var amount = parseFloat(body.match(/\$[\d\.]+\d/)[0].slice(1)) * -1
          if (UTILITIES[u].adjust) {
            let adjDate = new Date(UTILITIES[u].adjust.date);
            if (msg.getDate() > adjDate) {
              // Filter Main transactions for match
              let homeTrans = tMain.filter(r => Utilities.formatDate(new Date(r[0]), TZ, 'MM/DD/YYYY') == Utilities.formatDate(mDate, TZ, 'MM/DD/YYYY') && r[1] == u);
              amount = (homeTrans.length) ? amount - homeTrans[0][2] : amount * 0.65; // Approximation if not recorded
            }
          }

          let utility = (u.endsWith('Water')) ? u.replace('Water', 'Water District') : u;
          postEntry_([mDate, 5, utility, amount, 'Michael', UTILITIES[u].service]);
        } else if (UTILITIES[u].actMain) {
          // If Main account specified for utility...
          if (body.indexOf(UTILITIES[u].actMain)+1) {
            // If message includes Main account

            // Parse utility amount and attempt to post (will ignore duplicates)
            var amount = parseFloat(body.match(/\$[\d\.]+\d/)[0].slice(1)) * -1
            postMain_([mDate, u, amount, 'Michael', 'Utilities', UTILITIES[u].service]);
          } else {
            match = false;
          }
        } else {
          match = false;
        }
        if (!match) {
          // Unmatched for review
          Logger.log('No account match for message ' + m.getId() + ' - ' + m.getFirstMessageSubject());
          addTask(m.getId(), m.getFirstMessageSubject());
        }
      }
      catch (err) {
        Logger.log(err + ' for message ' + m.getId() + ' - ' + m.getFirstMessageSubject());
        m.addLabel(ERROR).moveToInbox();
      }
    }    
  }
}

/**
 * Function to post monthly rent collection based on spreadsheet configurations
 */
function getRent() {
  
  // authorize
  if (!auth_()) {
      return;
  }
  
  // Get total roommate utilities related to CRM
  var roomiePayment = SHEET.getRange(4, SHEET.getDataRange().getNumColumns()).getValue();
  
  // Get rent schedule from summary and post monthly entries
  var income = [];
  var rents = CRM.getSheetByName('Summary').getRange(4,6,2,3).getValues();
  rents.forEach(r => income.push([RENTDATE, 1, r[1], r[2], 'Michael', `${r[0]} ${MONTH} Rent`]));
  
  // Write income to transaction detail
  income.forEach(i => postEntry_(i));
  
}

/**
 * Function to post monthly mortgage payments based on spreadsheet configurations
 */
function getMortgage() {
  
  // authorize
  if (!auth_()) {
      return;
  }
  
  // Set mortgage amounts (Duplex)
  var mortgage = CRM.getSheetByName('Summary').getRange(8,6,2,4).getValues();
  mortgage = mortgage.map(m => [RENTDATE, m[1], m[2], m[3], 'Michael', m[0]]);
  
  // Write mortgage to transaction detail (Duplex)
  mortgage.forEach(m => postEntry_(m));

  // Set mortgage amounts (Main)
  var mortgageMain = SMAIN.getSheetByName('Summary').getRange(3,5,3,3).getValues();
  mortgageMain = mortgageMain.map(m => [RENTDATE, m[1], m[2], 'Michael', 'Mortgage', m[0]]);
  
  // Write mortgage to transaction detail (Main)
  mortgageMain.forEach(m => postMain_(m));
  
}

/**
 * Function to filldown summary formulas so pivots are accurate
 */
function closeBooks() {
  
  // authorize
  if (!auth_()) {
      return;
  }
  
  // Fill down transaction detail formalas for dates and accounts
  LEDGER.getRange(2, 1, 1, 3).autoFill(LEDGER.getRange(2, 1, LEDGER.getDataRange().getNumRows() - 1, 3), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  LEDGER.getRange(2, 11).autoFill(LEDGER.getRange(2, 11, LEDGER.getDataRange().getNumRows() - 1), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  // Fill down main home spreadsheet formulas
  TRANS.getRange(2, 8, 1, 2).autoFill(TRANS.getRange(2, 8, 2, 2), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  
}

/**
 * Function to prepare monthly email update for parents
 */
function sendUpdates(recipients) {
  
  // Get email template
  var html = HtmlService.createTemplateFromFile('Update');
  
  // Get data to populate and assign to template
  var data = {};
  data.dates = {
    'year': Utilities.formatDate(RENTDATE, TZ, 'yyyy'),
    'quarter': 'Q' + Math.ceil(((RENTDATE.getMonth() + 1) / 4))
  }
  data.cashflow = cashFlow_();
  data.finance = financing_();
  // API key stopped working... using Trulia scraping now
  data.zillow = getTrulia_(PROPS.getProperty('address_1'), PROPS.getProperty('city_state_zip'), PROPS.getProperty('trulia_id'));
  html.data = data;
  
  
  var images = getCharts_();

  html.images = Object.keys(images);

  // Send email update
  GmailApp.sendEmail(recipients, 
                      'Conan Rental Management ' + MONTH + ' Update', 
                      plainBody_(data), 
                      {
                        htmlBody: html.evaluate().getContent(), 
                        inlineImages: images, 
                        name: 'Conan Rental Management', 
                        cc: MAIN_USER
                      });
  
}

/**
 * Old function from roommate utility sharing
 */
function chargeUtilities() {
  
  // authorize
  if (!auth_()) {
      return;
  }
  
  var amounts = [];
  // Get latest bill from each utiltiy
  for (let u in UTILITIES) {
    
    try {
      // Get latest email by utility
      switch (UTILITIES[u].label) {
        case 'bills':
          var mail = BILLS.filter(m => m.getFirstMessageSubject().includes(UTILITIES[u].subject) && m.getMessages()[0].getFrom().indexOf(UTILITIES[u].from)+1)[0]
                    .getMessages()[0];
          break;
        case 'records':
          var mail = RECORDS.filter(m => m.getFirstMessageSubject().includes(UTILITIES[u].subject) && m.getMessages()[0].getFrom().indexOf(UTILITIES[u].from)+1)[0]
                    .getMessages()[0];
          break;
        default:
          Logger.log('none found');
      }
      
      var age = (new Date().getTime() - mail.getDate().getTime()) / (1000 * 60 * 60 * 24);
      if (age > 90) {
        GmailApp.sendEmail(MAIN_USER, 'CRM Script: Old Bill Added', 'Bill for ' + u + ' is ' + Math.ceil(age) + ' days old, please review.');
      }

      // Parse bill amount
      UTILITIES[u].amount = mail.getPlainBody().match(/\$[\d\.]+\d/)[0];
      amounts.push(UTILITIES[u].amount, '', mail.getFrom()+' - ['+Utilities.formatDate(mail.getDate(), TZ, "MM/dd/yyyy")+'] - '+mail.getSubject());
      }
      catch (err) {
        Logger.log(err, 'None found for '+ UTILITIES[u]);
        amounts.push('','','');
      }
    }
    
    // Write values to sheet
    var col = SHEET.getDataRange().getNumColumns()+1;
    SHEET.getRange(2, col).setValue(MONTH).setBorder(false, false, true, false, false, false);
    
    SHEET.getRange(5, col, amounts.length).setValues(amounts.map(a => [a]));
    
    // Get all formula cells
    var formulas = [3, 4, 20];
    for (let i = 6; i <= 21; i += 3) {
      formulas.push(i);
    }
  
  // Autofill formulas from prior month
  for (let f of formulas) {
    SHEET.getRange(f, col-1).autoFill(SHEET.getRange(f, col-1, 1, 2), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES)
  }
  
  // Send comment once completed - skipping as no longer duplex roomies
  //comment_(SS, [3, col]);
  
}

/**
 * Helper to save down bill attachment and extract PDF text
 */
function getAttachmentBody_(message) {
  let billFile = UTILITY_FOLD.createFile(message.getAttachments()[0].copyBlob()).getBlob();
  let newFile = Drive.Files.create({
    title: billFile.getName(),
    mimeType: 'application/vnd.google-apps.document',
  }, billFile);

  let text = DocumentApp.openById(newFile.id).getBody().getText();
  Drive.Files.remove(newFile.id);

  return text;
}

/**
 * Function to generate plain body version of email
 */
function plainBody_(data) {
  var body = PROPS.getProperty('address_1') + ' Update\n\n';
  for (let category in data) {
    body += '------------------------------\n' + category + '\n------------------------------\n';
    for (let value in data[category]) {
      body += '| ' + value + '\t|\t' + data[category][value] + ' |\n';
    }
  }

  return body;
}

/**
 * Function to get all charts in anaysis tab as images
 */
function getCharts_() {
  
  var images = {};
  
  // Get charts as blob images, assign to object
  var charts = CRM.getSheetByName('Analysis').getCharts();
  let pres = SlidesApp.create("temp");
  for (let i=0; i< charts.length; i++) {
    let chart = pres.getSlides()[0].insertSheetsChartAsImage(charts[i]);
    var blob = chart.getBlob().setName('chart'+i+'Blob');
    images['chart'+i] = blob;
    chart.remove();
  }

  DriveApp.getFileById(pres.getId()).setTrashed(true);
  
  return images;
}

/**
 * Function to grab cash flow stats from financials
 */
function cashFlow_() {
  
  // Get latest quarter, year, inception-to-date cash flow
  var flow = {};
  var financials = CRM.getSheetByName('Financials').getDataRange().getValues().slice(11)[0].slice(-3);
  ['QTD','YTD','ITD'].forEach((p, i) => flow[p] = Math.round(financials[i]).toLocaleString());
  
  return flow;
}

/**
 * Function to grab financing stats from summary and reimbursement
 */
function financing_() {
  
  var finance = {};
  
  // Get financing amounts from spreadsheet
  finance.mortgage = Math.round(CRM.getSheetByName('Summary').getRange(15, 2).getValue()).toLocaleString();
  finance.reimbursement = Math.round(CRM.getSheetByName('Reimbursement').getRange(17, 3).getValue()).toLocaleString();
  
  return finance;
}

/**
 * Custom spreadsheet function to summarize information from Trulia
 */
function TRULIASUMMARY(street, cityStateZip, tId) {
  var trulia = CacheService.getScriptCache().get('trulia');
  var params = CacheService.getScriptCache().get('params');
  if (trulia != null & params != null) {
    if ([street, cityStateZip] == params['address']) {
      return [['Trulia Estimate', trulia.price],
          ['Comparables', trulia.comp]];
    }
  }
  var trulia = getTrulia_(street, cityStateZip, tId);
  CacheService.getScriptCache().put('trulia', trulia, 21600);
  CacheService.getScriptCache().put('params', {'address': [street, cityStateZip]}, 21600);
    
  return [['Trulia Estimate', trulia.price],
          ['Comparables', trulia.comp]];
}

/**
 * Web scrape trulia property estimate and comparables
 */
function getTrulia_(stAddress, cityStateZip, tId) {
  let [city, state, zip] = cityStateZip.replace(',','').split(' ');
  let url = `https://www.trulia.com/p/${state}/${city}/${stAddress.replace(' ','-')}-${city}-${state}-${zip}--${tId}`
  let page = UrlFetchApp.fetch(url).getContentText();
  let price = page.match(/\>(\$[\d,]+)\</);
  price = parseInt(price[1].replace(/\W/g,''));

  let comps = [...page.matchAll(/\>(\$[\d,]*\d{3},\d{3})\</g)].map(m => m[1].replace(/\W/g,'')).filter(p => p != price);
  var avgComp = 0;
  comps.forEach(c => avgComp += parseInt(c));
  avgComp = Math.round(avgComp / comps.length);

  let data = {
    link: url,
    price: price.toLocaleString(),
    comp: avgComp.toLocaleString(),
    appreciation: (price - PROPS.getProperty('purchase_price')).toLocaleString()
  }

  return data;
}

/*
 * Checks for duplicate entry posting, then adds transaction to the spreadsheet
 * @param {array} entry - date, account, vendor, amount, reimbursement, purpose
 *
 */
function postEntry_(entry) {
  // Get all transactions in the log
  var transactions = LEDGER.getDataRange().getValues().slice(1).map(t => JSON.stringify([Utilities.formatDate(new Date(t[3]), TZ, "MM/dd/yyyy"), t[5], t[6]]));
  
  var posting = [Utilities.formatDate(new Date(entry[0]), TZ, "MM/dd/yyyy")].concat(entry.slice(1));
  if (transactions.indexOf(JSON.stringify([posting[0], posting[2], posting[3]])) == -1) {
    LEDGER.appendRow(['','',''].concat(posting));
    Logger.log(posting.slice(-1) + ' posted (Duplex) for ' + posting[0]);
  } else {
    Logger.log(posting.slice(-1) + ' already posted (Duplex) for ' + posting[0]);
  }
}

/*
 * Checks for duplicate entry posting, then adds transaction to the Main Home spreadsheet
 * @param {array} entry - date, vendor, amount, reimbursement, category, notes
 *
 */
function postMain_(entry) {
  var transactions = TRANS.getDataRange().getValues().slice(1).map(t => JSON.stringify([Utilities.formatDate(new Date(t[0]), TZ, "MM/dd/yyyy"), t[1], t[2]]));
  var posting = [Utilities.formatDate(new Date(entry[0]), TZ, "MM/dd/yyyy")].concat(entry.slice(1));
  if (transactions.indexOf(JSON.stringify(posting.slice(0,3))) == -1) {
    TRANS.appendRow(posting);
    Logger.log(posting.slice(-1) + ' posted (Main Home) for ' + posting[0]);
  } else {
    Logger.log(posting.slice(-1) + ' already posted (Main Home) for ' + posting[0]);
  }
}

function auth_() {
  // Nobody else gets to run this...
  if (MAIN_USER != PROPS.getProperty('main_user')) {
    SpreadsheetApp.getActive().toast(MAIN_USER + " - Don't use this script plz");
    return false;
  } else {
    return true;
  }
}

function addTask(msgId, subject) {
  let today = new Date();
  let tomorrow = new Date(today.getTime() + (1000 * 60 * 60 * 24));
  let task = {
      title: "Review Bill: " + subject,
      status:"needsAction",
      due: tomorrow.toISOString(),
      notes: "https://mail.google.com/mail/#all/" + msgId
    }

  try {
    // Call insert method with taskDetails and taskListId to insert Task to specified tasklist.
    task = Tasks.Tasks.insert(task, PROPS.getProperty('task_list'));
    // Print the Task ID of created task.
    console.log('Task with ID "%s" was created.', task.id);
  } catch (err) {
    // TODO (developer) - Handle exception from Tasks.insert() of Task API
    console.log('Failed with an error %s', err.message);
  }
}

// Currently not attaching to specific cell...
function comment_(spreadsheet, cell) {
  var fileId = spreadsheet.getId();
  var comment = {
    'content': `@${PROPS.getProperty('roomie_email')} ready for you :)`,
    'anchor': {
      'r': 'head', 
      'a': [
        {
          'page': 
          {
            'p': 0,
            'mp': 1
          },
          'matrix': 
          {
            'c': cell[1], 
            'r': cell[0]
          }
        }]
    }
  }
  Drive.Comments.insert(comment, fileId); 
}


