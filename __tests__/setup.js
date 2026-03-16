/**
 * Global mock setup for Google Apps Script services.
 * Loaded via Jest's setupFiles before each test module.
 *
 * Each service is a jest.fn()-based mock that mirrors the APIs used
 * by GmailService.js and Code.js so tests can run in Node.js without
 * a real Apps Script environment.
 */

// ── Utilities ─────────────────────────────────────────────────────────────────
global.Utilities = {
  formatDate: jest.fn((date, tz, fmt) => {
    // Return a deterministic string so duplicate-detection tests are predictable
    if (fmt === 'MM/dd/yyyy' || fmt === 'MM/DD/YYYY') return '01/01/2024';
    if (fmt === 'yyyy') return '2024';
    return String(date);
  }),
  base64Encode: jest.fn(data => Buffer.from(data).toString('base64')),
  base64EncodeWebSafe: jest.fn(data =>
    Buffer.from(typeof data === 'string' ? data : Buffer.from(data))
      .toString('base64')
      .replace(/\+/g, '-')
      .replace(/\//g, '_')
  ),
  base64DecodeWebSafe: jest.fn(str => {
    const normalized = str.replace(/-/g, '+').replace(/_/g, '/');
    return Array.from(Buffer.from(normalized, 'base64'));
  }),
  newBlob: jest.fn((data, mimeType, name) => ({
    _data: data,
    _mimeType: mimeType || 'application/octet-stream',
    _name: name || '',
    getBytes: jest.fn(function () { return Array.isArray(this._data) ? this._data : Array.from(Buffer.from(this._data || '')); }),
    getContentType: jest.fn(function () { return this._mimeType; }),
    getName: jest.fn(function () { return this._name; }),
    getDataAsString: jest.fn(function () {
      if (Array.isArray(this._data)) return Buffer.from(this._data).toString();
      return String(this._data || '');
    }),
    setName: jest.fn(function (n) { this._name = n; return this; })
  })),
  getUuid: jest.fn(() => 'test-uuid-1234')
};

// ── Logger ────────────────────────────────────────────────────────────────────
global.Logger = { log: jest.fn() };

// ── Session ───────────────────────────────────────────────────────────────────
global.Session = {
  getEffectiveUser: jest.fn(() => ({ getEmail: jest.fn(() => 'owner@example.com') }))
};

// ── PropertiesService ─────────────────────────────────────────────────────────
const _props = {
  main_user: 'owner@example.com',
  recipients: 'recipient@example.com',
  address_1: '123 Main St',
  city_state_zip: 'Portland OR 97201',
  trulia_id: 'tid123',
  purchase_price: '390000',
  task_list: 'tasklist_id_1'
};
global.PropertiesService = {
  getScriptProperties: jest.fn(() => ({
    getProperty: jest.fn(key => _props[key] || null),
    setProperty: jest.fn()
  }))
};

// ── SpreadsheetApp ────────────────────────────────────────────────────────────
const _makeSheet = (name, data = []) => ({
  _name: name,
  _data: data,
  getName: jest.fn(() => name),
  getDataRange: jest.fn(() => ({
    getValues: jest.fn(() => data),
    getNumRows: jest.fn(() => data.length),
    getNumColumns: jest.fn(() => (data[0] || []).length)
  })),
  getRange: jest.fn(() => ({
    getValues: jest.fn(() => [[]]),
    getValue: jest.fn(() => 0),
    setValue: jest.fn(),
    setValues: jest.fn(),
    setBorder: jest.fn(),
    autoFill: jest.fn(),
    clear: jest.fn()
  })),
  appendRow: jest.fn(),
  getCharts: jest.fn(() => [])
});

const _makeSpreadsheet = (sheets = {}) => ({
  getSheetByName: jest.fn(name => sheets[name] || _makeSheet(name)),
  getSpreadsheetTimeZone: jest.fn(() => 'America/Los_Angeles'),
  getId: jest.fn(() => 'spreadsheet_id_1'),
  toast: jest.fn()
});

global.SpreadsheetApp = {
  getActive: jest.fn(() => _makeSpreadsheet()),
  openById: jest.fn(() => _makeSpreadsheet()),
  AutoFillSeries: { DEFAULT_SERIES: 'DEFAULT_SERIES' }
};

// ── DriveApp ──────────────────────────────────────────────────────────────────
global.DriveApp = {
  getFileById: jest.fn(() => ({
    getBlob: jest.fn(() => ({
      getDataAsString: jest.fn(() => JSON.stringify({
        'Portland General Electric': {
          subject: 'Your PGE Bill',
          from: 'noreply@pge.com',
          account: 'ACC123',
          label: 'records',
          html: true,
          service: 'Electricity'
        }
      }))
    })),
    setTrashed: jest.fn()
  })),
  getFolderById: jest.fn(() => ({
    createFile: jest.fn(() => ({
      getBlob: jest.fn(() => ({ getName: jest.fn(() => 'bill.pdf') }))
    }))
  }))
};

// ── Drive (advanced service) ──────────────────────────────────────────────────
global.Drive = {
  Files: {
    create: jest.fn(() => ({ id: 'new_doc_id' })),
    remove: jest.fn()
  },
  Comments: {
    insert: jest.fn()
  }
};

// ── DocumentApp ───────────────────────────────────────────────────────────────
global.DocumentApp = {
  openById: jest.fn(() => ({
    getBody: jest.fn(() => ({
      getText: jest.fn(() => 'Due Date: 02/15/24\n$123.45\nACC123')
    }))
  }))
};

// ── Gmail (advanced service) ──────────────────────────────────────────────────
global.Gmail = {
  Users: {
    Labels: {
      list: jest.fn(() => ({
        labels: [
          { id: 'Label_1', name: 'Home/Bills' },
          { id: 'Label_2', name: 'Home/Duplex Records' },
          { id: 'Label_3', name: 'Script/Unmatched' },
          { id: 'Label_4', name: 'Script/Error' }
        ]
      }))
    },
    Threads: {
      list: jest.fn(() => ({ threads: [] })),
      get: jest.fn(() => ({ id: 'thread_1', messages: [] })),
      modify: jest.fn()
    },
    Messages: {
      get: jest.fn(() => ({ id: 'msg_1', payload: { headers: [], parts: [] } })),
      send: jest.fn(),
      Attachments: {
        get: jest.fn(() => ({ data: '' }))
      }
    }
  }
};

// ── GmailApp (standard service — used only for sendEmail fallback in tests) ───
global.GmailApp = {
  sendEmail: jest.fn(),
  getUserLabelByName: jest.fn(() => ({
    getThreads: jest.fn(() => []),
    addLabel: jest.fn()
  }))
};

// ── Tasks (advanced service) ──────────────────────────────────────────────────
global.Tasks = {
  Tasks: {
    insert: jest.fn(() => ({ id: 'task_id_1' }))
  }
};

// ── SlidesApp ─────────────────────────────────────────────────────────────────
global.SlidesApp = {
  create: jest.fn(() => ({
    getSlides: jest.fn(() => [{
      insertSheetsChartAsImage: jest.fn(() => ({
        getBlob: jest.fn(() => ({
          setName: jest.fn(function (n) { this._name = n; return this; }),
          getContentType: jest.fn(() => 'image/png'),
          getBytes: jest.fn(() => [])
        })),
        remove: jest.fn()
      }))
    }]),
    getId: jest.fn(() => 'pres_id_1')
  }))
};

// ── HtmlService ───────────────────────────────────────────────────────────────
global.HtmlService = {
  createTemplateFromFile: jest.fn(() => ({
    evaluate: jest.fn(() => ({ getContent: jest.fn(() => '<html>Update</html>') }))
  }))
};

// ── CacheService ──────────────────────────────────────────────────────────────
global.CacheService = {
  getScriptCache: jest.fn(() => ({
    get: jest.fn(() => null),
    put: jest.fn()
  }))
};

// ── UrlFetchApp ───────────────────────────────────────────────────────────────
global.UrlFetchApp = {
  fetch: jest.fn(() => ({
    getContentText: jest.fn(() => '<html><div>$450,000</div></html>')
  }))
};

// ── Console (already global in Node, ensure it works) ─────────────────────────
global.console = {
  log: jest.fn(),
  error: jest.fn()
};
