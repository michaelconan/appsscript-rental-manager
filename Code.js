/*
 * @module Conan Rental Management Automation
 * @author Michael Conan
 * @version 3.0
 * @date 2024-01-01
 * @description Scripts to pull utility bills from email, post recurring transactions and send
 *              monthly update for main home and Duplex rental property. Handles expense
 *              submissions from Google Form. Dependent on associated Google Sheets (per property)
 *              and a utility configuration JSON file stored in Drive.
 *
 *              Gmail operations use the Gmail advanced service (REST API) for improved performance.
 */

// ~~ GLOBALS ~~
// Wrapped in try/catch because simple triggers (onEdit) run with limited auth scope
try {
  // Testing flag — set to true in Tests.js to redirect emails to MAIN_USER
  var TEST = false;

  var PROPS = PropertiesService.getScriptProperties();
  var MAIN_USER = Session.getEffectiveUser().getEmail();
  var RECIPIENTS = PROPS.getProperty('recipients');

  // Utility bill configuration (JSON file in Drive)
  var UTILITIES = JSON.parse(DriveApp.getFileById('1j9y-WgXFvDztALYFMZr68WnLTapuhXyE').getBlob().getDataAsString());
  var UTILITY_FOLD = DriveApp.getFolderById('1GqO-9jl-vX4JSD00jO3OmbVIjyhB56ss');

  // CRM (Duplex) spreadsheet — active spreadsheet bound to this project
  var CRM = SpreadsheetApp.getActive();
  var LEDGER = CRM.getSheetByName('Transaction Detail');

  // Date/time config
  var TZ = CRM.getSpreadsheetTimeZone();
  var MONTH = Utilities.formatDate(new Date(), TZ, 'yyyy MMMM');
  var RENTDATE = new Date(new Date().getTime() + 4 * 24 * 60 * 60000);

  // Main Home spreadsheet
  var SMAIN = SpreadsheetApp.openById('1JwiQPOouopVfoE4wB0_lnGbq-8u-xBTDTNAEhR-AMJw');
  var TRANS = SMAIN.getSheetByName('Transactions');

} catch (e) {
  Logger.log('Some globals could not be loaded (limited auth context): ' + e);
}

// ── TransactionService ────────────────────────────────────────────────────────

/**
 * Handles duplicate-safe transaction posting to the Duplex and Main Home spreadsheets.
 */
class TransactionService {
  /**
   * @param {GoogleAppsScript.Spreadsheet.Sheet} ledger - Duplex Transaction Detail sheet
   * @param {GoogleAppsScript.Spreadsheet.Sheet} trans  - Main Home Transactions sheet
   * @param {string} tz - Spreadsheet timezone string
   */
  constructor(ledger, trans, tz) {
    this.ledger = ledger;
    this.trans = trans;
    this.tz = tz;
  }

  /**
   * Append a transaction to the Duplex ledger if it hasn't been posted before.
   * Duplicate detection is based on (date, vendor, amount).
   * @param {Array} entry - [date, account, vendor, amount, reimbursement, purpose]
   */
  postEntry(entry) {
    const transactions = this.ledger.getDataRange().getValues().slice(1)
      .map(t => JSON.stringify([
        Utilities.formatDate(new Date(t[3]), this.tz, 'MM/dd/yyyy'),
        t[5],
        t[6]
      ]));

    const posting = [Utilities.formatDate(new Date(entry[0]), this.tz, 'MM/dd/yyyy')].concat(entry.slice(1));
    const key = JSON.stringify([posting[0], posting[2], posting[3]]);

    if (!transactions.includes(key)) {
      this.ledger.appendRow(['', '', ''].concat(posting));
      Logger.log(posting.slice(-1) + ' posted (Duplex) for ' + posting[0]);
    } else {
      Logger.log(posting.slice(-1) + ' already posted (Duplex) for ' + posting[0]);
    }
  }

  /**
   * Append a transaction to the Main Home sheet if it hasn't been posted before.
   * Duplicate detection is based on (date, vendor, amount).
   * @param {Array} entry - [date, vendor, amount, reimbursement, category, notes]
   */
  postMain(entry) {
    const transactions = this.trans.getDataRange().getValues().slice(1)
      .map(t => JSON.stringify([
        Utilities.formatDate(new Date(t[0]), this.tz, 'MM/dd/yyyy'),
        t[1],
        t[2]
      ]));

    const posting = [Utilities.formatDate(new Date(entry[0]), this.tz, 'MM/dd/yyyy')].concat(entry.slice(1));
    const key = JSON.stringify(posting.slice(0, 3));

    if (!transactions.includes(key)) {
      this.trans.appendRow(posting);
      Logger.log(posting.slice(-1) + ' posted (Main Home) for ' + posting[0]);
    } else {
      Logger.log(posting.slice(-1) + ' already posted (Main Home) for ' + posting[0]);
    }
  }
}

// ── BillProcessor ─────────────────────────────────────────────────────────────

/**
 * Iterates configured utility accounts, finds matching Gmail threads, parses bill
 * amounts and dates, and posts transactions via TransactionService.
 */
class BillProcessor {
  /**
   * @param {GmailService} gmailService - Gmail API wrapper
   * @param {TransactionService} transactionService - Transaction poster
   * @param {Object} utilities - Utility config map keyed by vendor name
   * @param {GoogleAppsScript.Drive.Folder} utilityFolder - Drive folder for temp PDF files
   * @param {string} tz - Timezone string for date formatting
   * @param {GoogleAppsScript.Spreadsheet.Sheet} mainTrans - Main Home sheet for split adjustments
   */
  constructor(gmailService, transactionService, utilities, utilityFolder, tz, mainTrans) {
    this.gmailService = gmailService;
    this.transactionService = transactionService;
    this.utilities = utilities;
    this.utilityFolder = utilityFolder;
    this.tz = tz;
    this.mainTrans = mainTrans;
  }

  /**
   * Process all (or a scoped subset of) configured utility accounts:
   * find matching email threads, parse amounts, post to appropriate ledgers,
   * and create review tasks for unmatched bills.
   * @param {string[]} [scope] - Optional array of utility keys to process; processes all if omitted
   */
  processBills(scope) {
    const tMain = this.mainTrans.getDataRange().getValues();

    const billThreads = this.gmailService.getThreads('Home/Bills');
    const recordThreads = this.gmailService.getThreads('Home/Duplex Records');

    for (const u in this.utilities) {
      if (scope && !scope.includes(u)) continue;

      const utilConfig = this.utilities[u];
      const matchesThread = t =>
        this.gmailService.getThreadSubject(t).includes(utilConfig.subject) &&
        this.gmailService.getThreadFrom(t).includes(utilConfig.from);

      const threads = [...recordThreads.filter(matchesThread), ...billThreads.filter(matchesThread)];
      Logger.log(threads.length + ' bill threads for ' + u);

      for (const thread of threads) {
        let matched = true;
        const firstMsg = this.gmailService.getFirstMessage(thread);
        const msgId = this.gmailService.getMessageId(firstMsg);
        const threadId = this.gmailService.getThreadId(thread);

        try {
          const body = this._getMessageBody(firstMsg, msgId, utilConfig);

          // Use message date for payment confirmations, otherwise parse due date from body
          const mDate = utilConfig.payment
            ? this.gmailService.getMessageDate(firstMsg)
            : body.match(/Due Date.*(\d{2}\/\d{2}\/\d{2})/)[1];

          if (body.includes(utilConfig.account)) {
            // Bills-labeled entries are informational only — skip CRM posting
            if (utilConfig.label === 'bills') continue;

            let amount = parseFloat(body.match(/\$[\d\.]+\d/)[0].slice(1)) * -1;

            // Apply split adjustment if configured (e.g. shared utility after a date)
            if (utilConfig.adjust) {
              const adjDate = new Date(utilConfig.adjust.date);
              if (this.gmailService.getMessageDate(firstMsg) > adjDate) {
                const msgDate = this.gmailService.getMessageDate(firstMsg);
                const homeTrans = tMain.filter(r =>
                  Utilities.formatDate(new Date(r[0]), this.tz, 'MM/DD/YYYY') ===
                  Utilities.formatDate(msgDate, this.tz, 'MM/DD/YYYY') && r[1] === u
                );
                amount = homeTrans.length ? amount - homeTrans[0][2] : amount * 0.65;
              }
            }

            const vendor = u.endsWith('Water') ? u.replace('Water', 'Water District') : u;
            this.transactionService.postEntry([mDate, 5, vendor, amount, 'Michael', utilConfig.service]);

          } else if (utilConfig.actMain && body.includes(utilConfig.actMain)) {
            // Bill covers Main Home account
            const amount = parseFloat(body.match(/\$[\d\.]+\d/)[0].slice(1)) * -1;
            this.transactionService.postMain([mDate, u, amount, 'Michael', 'Utilities', utilConfig.service]);

          } else {
            matched = false;
          }

          if (!matched) {
            Logger.log('No account match for thread ' + threadId + ' - ' + this.gmailService.getThreadSubject(thread));
            addTask(msgId, this.gmailService.getThreadSubject(thread));
          }

        } catch (err) {
          Logger.log(err + ' for thread ' + threadId);
          const errorLabelId = this.gmailService.getLabelId('Script/Error');
          this.gmailService.addLabelToThread(threadId, errorLabelId);
          this.gmailService.moveThreadToInbox(threadId);
        }
      }
    }
  }

  /**
   * Extract the body text from a message according to the utility's config.
   * Supports PDF attachment, HTML body, or plain text.
   * @param {Object} message - Gmail API message object
   * @param {string} msgId - Message ID (needed for attachment download)
   * @param {Object} utilConfig - Utility configuration entry
   * @returns {string} Body text
   * @private
   */
  _getMessageBody(message, msgId, utilConfig) {
    if (utilConfig.attachment) {
      return this.gmailService.getAttachmentBody(message, msgId, this.utilityFolder);
    }
    return utilConfig.html
      ? this.gmailService.getHtmlBody(message)
      : this.gmailService.getPlainBody(message);
  }
}

// ── PropertyManager ───────────────────────────────────────────────────────────

/**
 * Orchestrates monthly property management: rent collection, mortgage recording,
 * formula maintenance, and email reporting.
 */
class PropertyManager {
  /**
   * @param {Object} config
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} config.crm  - Active CRM spreadsheet
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} config.smain - Main Home spreadsheet
   * @param {GoogleAppsScript.Properties.Properties} config.props   - Script properties
   * @param {string} config.tz         - Timezone string
   * @param {string} config.month      - Formatted month string (e.g. '2024 January')
   * @param {Date}   config.rentDate   - Date used for rent/mortgage entries (4 days out)
   * @param {GmailService}      config.gmailService       - Gmail API wrapper
   * @param {TransactionService} config.transactionService - Transaction poster
   */
  constructor({ crm, smain, props, tz, month, rentDate, gmailService, transactionService }) {
    this.crm = crm;
    this.smain = smain;
    this.props = props;
    this.tz = tz;
    this.month = month;
    this.rentDate = rentDate;
    this.gmailService = gmailService;
    this.transactionService = transactionService;
  }

  /**
   * Read rent schedule from the CRM Summary tab and post monthly income entries.
   */
  getRent() {
    const rents = this.crm.getSheetByName('Summary').getRange(4, 6, 2, 3).getValues();
    rents.forEach(r =>
      this.transactionService.postEntry([this.rentDate, 1, r[1], r[2], 'Michael', `${r[0]} ${this.month} Rent`])
    );
  }

  /**
   * Read mortgage schedules from both CRM and Main Home summaries and post payment entries.
   */
  getMortgage() {
    // Duplex mortgages
    let mortgage = this.crm.getSheetByName('Summary').getRange(8, 6, 2, 4).getValues();
    mortgage = mortgage.map(m => [this.rentDate, m[1], m[2], m[3], 'Michael', m[0]]);
    mortgage.forEach(m => this.transactionService.postEntry(m));

    // Main Home mortgages
    let mortgageMain = this.smain.getSheetByName('Summary').getRange(3, 5, 3, 3).getValues();
    mortgageMain = mortgageMain.map(m => [this.rentDate, m[1], m[2], 'Michael', 'Mortgage', m[0]]);
    mortgageMain.forEach(m => this.transactionService.postMain(m));
  }

  /**
   * Auto-fill formula columns in both Transaction Detail sheets so pivot tables stay current.
   */
  closeBooks() {
    const ledger = this.crm.getSheetByName('Transaction Detail');
    const ledgerRows = ledger.getDataRange().getNumRows() - 1;
    ledger.getRange(2, 1, 1, 3).autoFill(
      ledger.getRange(2, 1, ledgerRows, 3),
      SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES
    );
    ledger.getRange(2, 11).autoFill(
      ledger.getRange(2, 11, ledgerRows),
      SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES
    );

    const trans = this.smain.getSheetByName('Transactions');
    trans.getRange(2, 8, 1, 2).autoFill(
      trans.getRange(2, 8, 2, 2),
      SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES
    );
  }

  /**
   * Generate and send the monthly HTML email update to recipients.
   * @param {string} recipients - Comma-separated recipient email addresses
   */
  sendUpdates(recipients) {
    const html = HtmlService.createTemplateFromFile('Update');

    const data = {
      dates: {
        year: Utilities.formatDate(this.rentDate, this.tz, 'yyyy'),
        quarter: 'Q' + Math.ceil((this.rentDate.getMonth() + 1) / 4)
      },
      cashflow: this._cashFlow(),
      finance: this._financing(),
      zillow: getTrulia_(
        this.props.getProperty('address_1'),
        this.props.getProperty('city_state_zip'),
        this.props.getProperty('trulia_id')
      )
    };
    html.data = data;

    const images = this._getCharts();
    html.images = Object.keys(images);

    this.gmailService.sendEmail(
      recipients,
      'Conan Rental Management ' + this.month + ' Update',
      this._plainBody(data),
      {
        htmlBody: html.evaluate().getContent(),
        inlineImages: images,
        name: 'Conan Rental Management',
        cc: Session.getEffectiveUser().getEmail()
      }
    );
  }

  /**
   * Read QTD/YTD/ITD cash flow figures from the Financials sheet.
   * @returns {{ QTD: string, YTD: string, ITD: string }}
   */
  _cashFlow() {
    const flow = {};
    const financials = this.crm.getSheetByName('Financials').getDataRange().getValues().slice(11)[0].slice(-3);
    ['QTD', 'YTD', 'ITD'].forEach((p, i) => {
      flow[p] = Math.round(financials[i]).toLocaleString();
    });
    return flow;
  }

  /**
   * Read outstanding mortgage and reimbursement balances from the CRM.
   * @returns {{ mortgage: string, reimbursement: string }}
   */
  _financing() {
    return {
      mortgage: Math.round(this.crm.getSheetByName('Summary').getRange(15, 2).getValue()).toLocaleString(),
      reimbursement: Math.round(this.crm.getSheetByName('Reimbursement').getRange(17, 3).getValue()).toLocaleString()
    };
  }

  /**
   * Extract all charts from the Analysis sheet as PNG image blobs.
   * Uses a temporary Slides presentation as the rendering intermediary.
   * @returns {Object} Map of chart name -> Blob
   */
  _getCharts() {
    const images = {};
    const charts = this.crm.getSheetByName('Analysis').getCharts();
    const pres = SlidesApp.create('temp');
    for (let i = 0; i < charts.length; i++) {
      const chart = pres.getSlides()[0].insertSheetsChartAsImage(charts[i]);
      images['chart' + i] = chart.getBlob().setName('chart' + i + 'Blob');
      chart.remove();
    }
    DriveApp.getFileById(pres.getId()).setTrashed(true);
    return images;
  }

  /**
   * Build the plain-text fallback version of the monthly update email.
   * @param {Object} data - Email data object with nested category maps
   * @returns {string} Formatted plain text
   */
  _plainBody(data) {
    let body = this.props.getProperty('address_1') + ' Update\n\n';
    for (const category in data) {
      body += '------------------------------\n' + category + '\n------------------------------\n';
      for (const value in data[category]) {
        body += '| ' + value + '\t|\t' + data[category][value] + ' |\n';
      }
    }
    return body;
  }
}

// ── Entry point functions (called by triggers / manual execution) ─────────────

/**
 * Main monthly orchestration: collect rent, record mortgage, close books, send update.
 * getBills() is intentionally omitted — run it on a separate trigger to avoid timeout.
 */
function manageProperty() {
  if (!auth_()) return;

  const gmailService = new GmailService();
  const transService = new TransactionService(LEDGER, TRANS, TZ);
  const manager = new PropertyManager({
    crm: CRM, smain: SMAIN, props: PROPS,
    tz: TZ, month: MONTH, rentDate: RENTDATE,
    gmailService, transactionService: transService
  });

  manager.getRent();
  manager.getMortgage();
  manager.closeBooks();
  manager.sendUpdates(TEST ? MAIN_USER : RECIPIENTS);
}

/**
 * Parse utility bill emails and post transactions to ledgers.
 * Intended to run on a separate trigger from manageProperty() due to runtime limits.
 * @param {string[]} [scope] - Optional array of utility keys to limit processing
 */
function getBills(scope) {
  if (!auth_()) return;

  const gmailService = new GmailService();
  const transService = new TransactionService(LEDGER, TRANS, TZ);
  const processor = new BillProcessor(
    gmailService, transService, UTILITIES, UTILITY_FOLD, TZ, TRANS
  );
  processor.processBills(scope);
}

/**
 * Custom spreadsheet function: return Trulia estimate and comparables with 6-hour caching.
 * @param {string} street - Street address
 * @param {string} cityStateZip - City, state and zip (e.g. 'Portland OR 97201')
 * @param {string} tId - Trulia property ID
 * @returns {Array[][]} 2-row array: [[label, value], [label, value]]
 */
function TRULIASUMMARY(street, cityStateZip, tId) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('trulia');
  const cachedParams = cache.get('trulia_params');

  if (cached && cachedParams) {
    const params = JSON.parse(cachedParams);
    if (params.street === street && params.cityStateZip === cityStateZip) {
      const trulia = JSON.parse(cached);
      return [['Trulia Estimate', trulia.price], ['Comparables', trulia.comp]];
    }
  }

  const trulia = getTrulia_(street, cityStateZip, tId);
  cache.put('trulia', JSON.stringify(trulia), 21600);
  cache.put('trulia_params', JSON.stringify({ street, cityStateZip }), 21600);

  return [['Trulia Estimate', trulia.price], ['Comparables', trulia.comp]];
}

// ── Standalone helpers ────────────────────────────────────────────────────────

/**
 * Web-scrape Trulia for a property's current estimate and comparable sales.
 * @param {string} stAddress - Street address
 * @param {string} cityStateZip - 'City State Zip' string
 * @param {string} tId - Trulia listing ID
 * @returns {{ link: string, price: string, comp: string, appreciation: string }}
 */
function getTrulia_(stAddress, cityStateZip, tId) {
  const [city, state, zip] = cityStateZip.replace(',', '').split(' ');
  const url = `https://www.trulia.com/p/${state}/${city}/${stAddress.replace(' ', '-')}-${city}-${state}-${zip}--${tId}`;
  const page = UrlFetchApp.fetch(url).getContentText();

  let price = page.match(/\>(\$[\d,]+)\</);
  price = parseInt(price[1].replace(/\W/g, ''));

  const comps = [...page.matchAll(/\>(\$[\d,]*\d{3},\d{3})\</g)]
    .map(m => m[1].replace(/\W/g, ''))
    .filter(p => parseInt(p) !== price);

  let avgComp = 0;
  comps.forEach(c => avgComp += parseInt(c));
  avgComp = Math.round(avgComp / comps.length);

  return {
    link: url,
    price: price.toLocaleString(),
    comp: avgComp.toLocaleString(),
    appreciation: (price - parseInt(PROPS.getProperty('purchase_price'))).toLocaleString()
  };
}

/**
 * Verify that the current user is the configured main user; toast and return false otherwise.
 * @returns {boolean}
 */
function auth_() {
  if (MAIN_USER !== PROPS.getProperty('main_user')) {
    SpreadsheetApp.getActive().toast(MAIN_USER + " - Don't use this script plz");
    return false;
  }
  return true;
}

/**
 * Create a Google Task linked to a Gmail message for manual bill review.
 * @param {string} msgId - Gmail message ID
 * @param {string} subject - Message subject (used as task title)
 */
function addTask(msgId, subject) {
  const tomorrow = new Date(new Date().getTime() + 1000 * 60 * 60 * 24);
  const task = {
    title: 'Review Bill: ' + subject,
    status: 'needsAction',
    due: tomorrow.toISOString(),
    notes: 'https://mail.google.com/mail/#all/' + msgId
  };
  try {
    const created = Tasks.Tasks.insert(task, PROPS.getProperty('task_list'));
    console.log('Task with ID "%s" was created.', created.id);
  } catch (err) {
    console.log('Failed with an error %s', err.message);
  }
}

// Node.js / Jest compatibility — ignored by Apps Script runtime
if (typeof module !== 'undefined') {
  module.exports = { TransactionService, BillProcessor, PropertyManager, getTrulia_, auth_, addTask };
}
