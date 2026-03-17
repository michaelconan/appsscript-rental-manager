/**
 * Unit tests for BillProcessor and TransactionService.
 */

const { BillProcessor, TransactionService } = require('../Code');

// ── Helpers ───────────────────────────────────────────────────────────────────

const makeMessage = (overrides = {}) => ({
  id: 'msg_001',
  payload: { headers: [{ name: 'Date', value: 'Thu, 01 Feb 2024 10:00:00 +0000' }] },
  ...overrides
});

const makeThread = (msg = makeMessage(), overrides = {}) => ({
  id: 'thread_001',
  messages: [msg],
  ...overrides
});

// Build a mock GmailService whose behaviour can be customised per test
const makeMockGmailService = (bodyText = 'Due Date: 02/15/24\n$123.45\nACC123') => ({
  getThreads: jest.fn(labelName => {
    if (labelName === 'Home/Duplex Records') return [makeThread()];
    return [];
  }),
  getFirstMessage: jest.fn(t => t.messages[0]),
  getMessageId: jest.fn(msg => msg.id),
  getThreadId: jest.fn(t => t.id),
  getThreadSubject: jest.fn(() => 'Your PGE Bill'),
  getThreadFrom: jest.fn(() => 'noreply@pge.com'),
  getMessageDate: jest.fn(() => new Date('2024-02-01')),
  getPlainBody: jest.fn(() => bodyText),
  getHtmlBody: jest.fn(() => '<p>' + bodyText + '</p>'),
  getAttachmentBody: jest.fn(() => bodyText),
  getLabelId: jest.fn(name => ({ 'Script/Error': 'Label_4' }[name] || 'Label_X')),
  addLabelToThread: jest.fn(),
  moveThreadToInbox: jest.fn()
});

// Minimal utility config for a standard duplex bill
const makeUtilities = (overrides = {}) => ({
  'Portland General Electric': {
    subject: 'Your PGE Bill',
    from: 'noreply@pge.com',
    account: 'ACC123',
    label: 'records',
    html: false,
    service: 'Electricity',
    ...overrides
  }
});

// Build a mock sheet that looks like it already has one header row
const makeSheet = (rows = [['header']]) => ({
  getDataRange: jest.fn(() => ({
    getValues: jest.fn(() => rows),
    getNumRows: jest.fn(() => rows.length)
  })),
  appendRow: jest.fn()
});

// ── TransactionService ────────────────────────────────────────────────────────

describe('TransactionService', () => {
  let ledger, trans, svc;

  beforeEach(() => {
    jest.clearAllMocks();
    // Header row + one existing transaction: date col[3], vendor col[5], amount col[6]
    ledger = makeSheet([
      ['', '', '', '', '', '', ''],
      ['', '', '', '2024-01-01', '', 'PGE', -100]
    ]);
    trans = makeSheet([
      ['Date', 'Vendor', 'Amount'],  // header row consumed by slice(1)
      ['2024-01-01', 'PGE', -100],
      ['2024-01-05', 'Water', -50]
    ]);
    svc = new TransactionService(ledger, trans, 'America/Los_Angeles');
  });

  describe('postEntry', () => {
    it('appends a new Duplex transaction when not already posted', () => {
      Utilities.formatDate.mockImplementation((d, tz, fmt) => '02/01/2024');

      svc.postEntry([new Date('2024-02-01'), 5, 'PGE', -123.45, 'Michael', 'Electricity']);

      expect(ledger.appendRow).toHaveBeenCalledWith(
        ['', '', '', '02/01/2024', 5, 'PGE', -123.45, 'Michael', 'Electricity']
      );
    });

    it('skips posting when an identical entry already exists', () => {
      // Make Utilities.formatDate always return the same date string as in the existing row
      Utilities.formatDate.mockReturnValue('01/01/2024');

      svc.postEntry([new Date('2024-01-01'), 5, 'PGE', -100, 'Michael', 'Electricity']);

      expect(ledger.appendRow).not.toHaveBeenCalled();
    });
  });

  describe('postMain', () => {
    it('appends a new Main Home transaction when not already posted', () => {
      Utilities.formatDate.mockReturnValue('03/01/2024');

      svc.postMain([new Date('2024-03-01'), 'Water', -75, 'Michael', 'Utilities', 'Water bill']);

      expect(trans.appendRow).toHaveBeenCalledWith(
        ['03/01/2024', 'Water', -75, 'Michael', 'Utilities', 'Water bill']
      );
    });

    it('skips posting when a duplicate Main Home entry already exists', () => {
      Utilities.formatDate.mockReturnValue('01/01/2024');

      svc.postMain([new Date('2024-01-01'), 'PGE', -100, 'Michael', 'Utilities', 'PGE bill']);

      expect(trans.appendRow).not.toHaveBeenCalled();
    });
  });
});

// ── BillProcessor ─────────────────────────────────────────────────────────────

describe('BillProcessor', () => {
  let gmailSvc, transSvc, processor;

  beforeEach(() => {
    jest.clearAllMocks();
    gmailSvc = makeMockGmailService();
    transSvc = { postEntry: jest.fn(), postMain: jest.fn() };

    const mainTrans = makeSheet([
      ['Date', 'Vendor', 'Amount'],
      [new Date('2024-01-01'), 'PGE', -50]
    ]);

    processor = new BillProcessor(
      gmailSvc, transSvc,
      makeUtilities(),
      { createFile: jest.fn() }, // utilityFolder mock
      'America/Los_Angeles',
      mainTrans
    );
  });

  describe('processBills', () => {
    it('posts a Duplex entry when account number matches', () => {
      processor.processBills();
      expect(transSvc.postEntry).toHaveBeenCalledWith(
        expect.arrayContaining([-123.45])
      );
    });

    it('posts a Main Home entry when actMain account matches', () => {
      processor = new BillProcessor(
        gmailSvc, transSvc,
        makeUtilities({ account: 'OTHER', actMain: 'ACC123' }),
        {},
        'America/Los_Angeles',
        makeSheet()
      );
      processor.processBills();
      expect(transSvc.postMain).toHaveBeenCalled();
    });

    it('skips bills-label entries for CRM posting', () => {
      processor = new BillProcessor(
        gmailSvc, transSvc,
        makeUtilities({ label: 'bills' }),
        {},
        'America/Los_Angeles',
        makeSheet()
      );
      processor.processBills();
      expect(transSvc.postEntry).not.toHaveBeenCalled();
    });

    it('creates a task for threads where no account matches', () => {
      global.Tasks = { Tasks: { insert: jest.fn(() => ({ id: 'task_1' })) } };
      global.PROPS = PropertiesService.getScriptProperties();

      processor = new BillProcessor(
        gmailSvc, transSvc,
        makeUtilities({ account: 'NOMATCH', actMain: undefined }),
        {},
        'America/Los_Angeles',
        makeSheet()
      );
      processor.processBills();
      expect(transSvc.postEntry).not.toHaveBeenCalled();
    });

    it('adds Error label and moves to inbox when parsing throws', () => {
      gmailSvc.getPlainBody.mockImplementation(() => { throw new Error('parse error'); });
      processor.processBills();
      expect(gmailSvc.addLabelToThread).toHaveBeenCalledWith('thread_001', 'Label_4');
      expect(gmailSvc.moveThreadToInbox).toHaveBeenCalledWith('thread_001');
    });

    it('respects scope filtering — skips utilities not in scope', () => {
      processor.processBills(['SomeOtherUtility']);
      expect(transSvc.postEntry).not.toHaveBeenCalled();
    });

    it('applies the amount adjustment when adjust config is present and message is recent', () => {
      const mainTrans = makeSheet([
        ['Date', 'Vendor', 'Amount'],
        [new Date('2024-02-01'), 'Portland General Electric', -50]
      ]);
      Utilities.formatDate.mockReturnValue('02/01/2024');

      processor = new BillProcessor(
        gmailSvc, transSvc,
        makeUtilities({ adjust: { date: '2024-01-01' } }),
        {},
        'America/Los_Angeles',
        mainTrans
      );
      processor.processBills();
      // Amount should be adjusted: -123.45 - (-50) = -73.45
      const call = transSvc.postEntry.mock.calls[0][0];
      expect(call[3]).toBeCloseTo(-73.45, 1);
    });

    it('approximates adjustment (×0.65) when no matching Main Home transaction', () => {
      const mainTrans = makeSheet([['Date', 'Vendor', 'Amount']]);
      Utilities.formatDate.mockReturnValue('02/01/2024');

      processor = new BillProcessor(
        gmailSvc, transSvc,
        makeUtilities({ adjust: { date: '2024-01-01' } }),
        {},
        'America/Los_Angeles',
        mainTrans
      );
      processor.processBills();
      const call = transSvc.postEntry.mock.calls[0][0];
      expect(call[3]).toBeCloseTo(-123.45 * 0.65, 1);
    });

    it('appends "Water District" suffix when vendor ends with Water', () => {
      gmailSvc.getThreadSubject.mockReturnValue('Tualatin Valley Water Bill');
      gmailSvc.getThreadFrom.mockReturnValue('noreply@tvwd.com');
      gmailSvc.getPlainBody.mockReturnValue('Due Date: 02/15/24\n$55.00\nWATER99');

      processor = new BillProcessor(
        gmailSvc, transSvc,
        {
          'Tualatin Valley Water': {
            subject: 'Tualatin Valley Water Bill',
            from: 'noreply@tvwd.com',
            account: 'WATER99',
            label: 'records',
            html: false,
            service: 'Water'
          }
        },
        {},
        'America/Los_Angeles',
        makeSheet()
      );
      processor.processBills();
      const call = transSvc.postEntry.mock.calls[0][0];
      expect(call[2]).toBe('Tualatin Valley Water District');
    });
  });

  describe('_getMessageBody', () => {
    it('returns plain body for plain-text config', () => {
      const msg = makeMessage();
      const body = processor._getMessageBody(msg, 'msg_001', { html: false });
      expect(gmailSvc.getPlainBody).toHaveBeenCalledWith(msg);
      expect(body).toBe('Due Date: 02/15/24\n$123.45\nACC123');
    });

    it('returns HTML body for html config', () => {
      const msg = makeMessage();
      const body = processor._getMessageBody(msg, 'msg_001', { html: true });
      expect(gmailSvc.getHtmlBody).toHaveBeenCalledWith(msg);
    });

    it('calls getAttachmentBody for attachment config', () => {
      const msg = makeMessage();
      const folder = {};
      processor.utilityFolder = folder;
      processor._getMessageBody(msg, 'msg_001', { attachment: true });
      expect(gmailSvc.getAttachmentBody).toHaveBeenCalledWith(msg, 'msg_001', folder);
    });
  });
});
