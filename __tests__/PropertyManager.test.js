/**
 * Unit tests for PropertyManager.
 */

const { PropertyManager } = require('../Code');

// ── Helpers ───────────────────────────────────────────────────────────────────

const makeSheet = (name, returnValues = [[]]) => ({
  _name: name,
  getName: jest.fn(() => name),
  getDataRange: jest.fn(() => ({
    getValues: jest.fn(() => returnValues),
    getNumRows: jest.fn(() => returnValues.length)
  })),
  getRange: jest.fn(() => ({
    getValues: jest.fn(() => returnValues.slice(0, 2)),
    getValue: jest.fn(() => 500000),
    autoFill: jest.fn()
  })),
  getCharts: jest.fn(() => []),
  appendRow: jest.fn()
});

const makeSpreadsheet = (sheetMap = {}) => ({
  getSheetByName: jest.fn(name => sheetMap[name] || makeSheet(name)),
  getId: jest.fn(() => 'crm_id')
});

const makeTransactionService = () => ({
  postEntry: jest.fn(),
  postMain: jest.fn()
});

const makeGmailService = () => ({
  sendEmail: jest.fn()
});

const makeProps = (overrides = {}) => ({
  getProperty: jest.fn(key => ({
    address_1: '123 Main St',
    city_state_zip: 'Portland OR 97201',
    trulia_id: 'tid_1',
    purchase_price: '390000',
    ...overrides
  }[key] || null))
});

const rentDate = new Date('2024-01-05');

const makeManager = (crmSheets = {}, smainSheets = {}, overrides = {}) => {
  const crm = makeSpreadsheet(crmSheets);
  const smain = makeSpreadsheet(smainSheets);
  const transService = makeTransactionService();
  const gmailService = makeGmailService();
  const props = makeProps();

  const manager = new PropertyManager({
    crm, smain, props,
    tz: 'America/Los_Angeles',
    month: '2024 January',
    rentDate,
    gmailService,
    transactionService: transService,
    ...overrides
  });

  manager._transactionService = transService;
  manager._gmailService = gmailService;
  return manager;
};

// ── getRent ───────────────────────────────────────────────────────────────────

describe('PropertyManager.getRent', () => {
  it('reads Summary and posts one income entry per rent row', () => {
    const rentRows = [
      ['Unit A', 'Tenant A', 1500],
      ['Unit B', 'Tenant B', 1200]
    ];
    const summarySheet = {
      ...makeSheet('Summary'),
      getRange: jest.fn(() => ({ getValues: jest.fn(() => rentRows) }))
    };
    const manager = makeManager({ Summary: summarySheet });

    manager.getRent();

    expect(manager.transactionService.postEntry).toHaveBeenCalledTimes(2);
    expect(manager.transactionService.postEntry).toHaveBeenCalledWith(
      [rentDate, 1, 'Tenant A', 1500, 'Michael', 'Unit A 2024 January Rent']
    );
    expect(manager.transactionService.postEntry).toHaveBeenCalledWith(
      [rentDate, 1, 'Tenant B', 1200, 'Michael', 'Unit B 2024 January Rent']
    );
  });
});

// ── getMortgage ───────────────────────────────────────────────────────────────

describe('PropertyManager.getMortgage', () => {
  it('posts duplex mortgage entries from CRM Summary', () => {
    const duplexMortgageRows = [
      ['Duplex Mortgage', 3, 'Lender A', 2500],
      ['Duplex HELOC', 4, 'Lender B', 300]
    ];
    const mainMortgageRows = [
      ['Main Mortgage', 1, 2200],
      ['Main HELOC', 2, 150],
      ['Main PMI', 3, 50]
    ];

    const crmSummary = {
      ...makeSheet('Summary'),
      getRange: jest.fn(() => ({ getValues: jest.fn(() => duplexMortgageRows) }))
    };
    const smainSummary = {
      ...makeSheet('Summary'),
      getRange: jest.fn(() => ({ getValues: jest.fn(() => mainMortgageRows) }))
    };

    const manager = makeManager(
      { Summary: crmSummary },
      { Summary: smainSummary }
    );

    manager.getMortgage();

    // 2 duplex rows → postEntry twice
    expect(manager.transactionService.postEntry).toHaveBeenCalledTimes(2);
    // 3 main rows → postMain three times
    expect(manager.transactionService.postMain).toHaveBeenCalledTimes(3);
  });
});

// ── closeBooks ────────────────────────────────────────────────────────────────

describe('PropertyManager.closeBooks', () => {
  it('calls autoFill on the correct ranges for both spreadsheets', () => {
    const autoFill = jest.fn();
    const ledgerRange = { autoFill };
    const transRange = { autoFill };

    const ledgerSheet = {
      ...makeSheet('Transaction Detail'),
      getDataRange: jest.fn(() => ({ getNumRows: jest.fn(() => 10) })),
      getRange: jest.fn(() => ledgerRange)
    };
    const transSheet = {
      ...makeSheet('Transactions'),
      getRange: jest.fn(() => transRange)
    };

    const manager = makeManager(
      { 'Transaction Detail': ledgerSheet },
      { Transactions: transSheet }
    );

    manager.closeBooks();

    // autoFill called for: col 1-3, col 11 (ledger) + col 8-9 (trans)
    expect(autoFill).toHaveBeenCalledTimes(3);
  });
});

// ── _cashFlow ─────────────────────────────────────────────────────────────────

describe('PropertyManager._cashFlow', () => {
  it('returns QTD, YTD, ITD formatted from the Financials sheet', () => {
    // Row at index 11 (after slice(11)), last 3 values = [1000, 5000, 20000]
    const finData = Array(12).fill([]).map((_, i) =>
      i === 11 ? [0, 0, 0, 0, 0, 1000, 5000, 20000] : []
    );
    const financialsSheet = {
      ...makeSheet('Financials'),
      getDataRange: jest.fn(() => ({ getValues: jest.fn(() => finData) }))
    };

    const manager = makeManager({ Financials: financialsSheet });
    const result = manager._cashFlow();

    expect(result).toHaveProperty('QTD');
    expect(result).toHaveProperty('YTD');
    expect(result).toHaveProperty('ITD');
    expect(result.QTD).toBe('1,000');
    expect(result.YTD).toBe('5,000');
    expect(result.ITD).toBe('20,000');
  });
});

// ── _financing ────────────────────────────────────────────────────────────────

describe('PropertyManager._financing', () => {
  it('returns mortgage and reimbursement amounts from spreadsheet', () => {
    const summarySheet = {
      ...makeSheet('Summary'),
      getRange: jest.fn(() => ({ getValue: jest.fn(() => 250000.7) }))
    };
    const reimbursementSheet = {
      ...makeSheet('Reimbursement'),
      getRange: jest.fn(() => ({ getValue: jest.fn(() => 12345.3) }))
    };

    const manager = makeManager({
      Summary: summarySheet,
      Reimbursement: reimbursementSheet
    });
    const result = manager._financing();

    expect(result.mortgage).toBe('250,001');
    expect(result.reimbursement).toBe('12,345');
  });
});

// ── _plainBody ────────────────────────────────────────────────────────────────

describe('PropertyManager._plainBody', () => {
  it('formats all data categories and their values', () => {
    const manager = makeManager();
    const data = {
      cashflow: { QTD: '1,000', YTD: '5,000', ITD: '20,000' },
      finance: { mortgage: '250,001', reimbursement: '12,345' }
    };

    const body = manager._plainBody(data);

    expect(body).toContain('123 Main St Update');
    expect(body).toContain('cashflow');
    expect(body).toContain('QTD');
    expect(body).toContain('1,000');
    expect(body).toContain('finance');
    expect(body).toContain('mortgage');
    expect(body).toContain('250,001');
  });
});

// ── sendUpdates ───────────────────────────────────────────────────────────────

describe('PropertyManager.sendUpdates', () => {
  it('evaluates the HTML template and sends via gmailService', () => {
    // Financial sheets for sendUpdates internals
    const finData = Array(12).fill([]).map((_, i) =>
      i === 11 ? [0, 0, 0, 0, 0, 500, 2000, 10000] : []
    );
    const crmSheets = {
      Financials: { ...makeSheet('Financials'), getDataRange: jest.fn(() => ({ getValues: jest.fn(() => finData) })) },
      Summary: { ...makeSheet('Summary'), getRange: jest.fn(() => ({ getValue: jest.fn(() => 200000) })) },
      Reimbursement: { ...makeSheet('Reimbursement'), getRange: jest.fn(() => ({ getValue: jest.fn(() => 5000) })) },
      Analysis: { ...makeSheet('Analysis'), getCharts: jest.fn(() => []) }
    };

    // Mock UrlFetchApp for getTrulia_ (called inside sendUpdates)
    global.UrlFetchApp.fetch.mockReturnValue({
      getContentText: jest.fn(() => '>$450,000< >$460,000<')
    });

    const gmailService = makeGmailService();
    const manager = makeManager(crmSheets, {}, {
      gmailService,
      transactionService: makeTransactionService()
    });
    // Patch internal references to use the injected services
    manager.gmailService = gmailService;

    manager.sendUpdates('recipient@example.com');

    expect(HtmlService.createTemplateFromFile).toHaveBeenCalledWith('Update');
    expect(gmailService.sendEmail).toHaveBeenCalledWith(
      'recipient@example.com',
      expect.stringContaining('Conan Rental Management'),
      expect.any(String),
      expect.objectContaining({ name: 'Conan Rental Management' })
    );
  });
});
