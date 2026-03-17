/**
 * Apps Script integration test functions.
 * Run these manually from the Apps Script editor to verify live behaviour.
 * For unit tests with mocks, see the __tests__/ directory (Jest).
 */

/** Full monthly workflow — emails go to MAIN_USER instead of RECIPIENTS. */
function testManageProperty() {
  TEST = true;
  manageProperty();
}

/** Process all utility bill emails in test mode. */
function testGetBills() {
  TEST = true;
  getBills();
}

/** Process a single utility vendor to verify parsing. */
function testGetVendor() {
  TEST = true;
  getBills(['Tualatin Valley Water District']);
}

/** Simulate an edit event on the last row of the Transaction Detail sheet. */
function testEditTrigger() {
  const evt = { range: LEDGER.getRange(LEDGER.getDataRange().getNumRows(), 4) };
  onEdit(evt);
}

/** Simple connectivity check — returns 'Hello <message>'. */
function apiTest(message) {
  return 'Hello ' + message;
}

/** Create a test review task linked to a sample Gmail message. */
function testAddTask() {
  addTask('185eeb3794bb9001', 'test task');
}

/** Verify that financial data helpers return expected shape. */
function testFinancials() {
  const transService = new TransactionService(LEDGER, TRANS, TZ);
  const manager = new PropertyManager({
    crm: CRM, smain: SMAIN, props: PROPS,
    tz: TZ, month: MONTH, rentDate: RENTDATE,
    gmailService: new GmailService(),
    transactionService: transService
  });
  console.log('cashFlow:', manager._cashFlow());
  console.log('financing:', manager._financing());
}
