/**
 * Unit tests for GmailService — Gmail advanced service wrapper.
 */

const { GmailService } = require('../GmailService');

// Helper: build a base64url-encoded string the way the real API does
const b64 = str => Buffer.from(str).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');

// Factory for a minimal Gmail API message object
const makeMessage = (overrides = {}) => ({
  id: 'msg_001',
  threadId: 'thread_001',
  internalDate: '1700000000000',
  payload: {
    mimeType: 'text/plain',
    headers: [
      { name: 'Subject', value: 'Your Bill is Ready' },
      { name: 'From', value: 'billing@utility.com' },
      { name: 'Date', value: 'Thu, 01 Feb 2024 10:00:00 +0000' }
    ],
    body: { data: b64('Due Date: 02/15/24\n$123.45\nACC123') }
  },
  ...overrides
});

// Factory for a multipart message
const makeMultipartMessage = () => ({
  id: 'msg_002',
  internalDate: '1700000000000',
  payload: {
    mimeType: 'multipart/alternative',
    headers: [
      { name: 'Subject', value: 'Multipart Bill' },
      { name: 'From', value: 'info@provider.com' },
      { name: 'Date', value: 'Fri, 02 Feb 2024 08:00:00 +0000' }
    ],
    parts: [
      { mimeType: 'text/plain', body: { data: b64('Plain text body') } },
      { mimeType: 'text/html', body: { data: b64('<p>HTML body</p>') } }
    ]
  }
});

describe('GmailService', () => {
  let svc;

  beforeEach(() => {
    jest.clearAllMocks();
    svc = new GmailService();
  });

  // ── getLabelId ─────────────────────────────────────────────────────────────

  describe('getLabelId', () => {
    it('returns the matching label ID', () => {
      const id = svc.getLabelId('Home/Bills');
      expect(Gmail.Users.Labels.list).toHaveBeenCalledWith('me');
      expect(id).toBe('Label_1');
    });

    it('caches the label ID on subsequent calls', () => {
      svc.getLabelId('Home/Bills');
      svc.getLabelId('Home/Bills');
      expect(Gmail.Users.Labels.list).toHaveBeenCalledTimes(1);
    });

    it('throws if the label is not found', () => {
      expect(() => svc.getLabelId('Nonexistent/Label')).toThrow('Gmail label not found: Nonexistent/Label');
    });
  });

  // ── getThreads ─────────────────────────────────────────────────────────────

  describe('getThreads', () => {
    it('lists threads by label ID and fetches each full thread', () => {
      Gmail.Users.Threads.list.mockReturnValue({
        threads: [{ id: 'thread_A' }, { id: 'thread_B' }]
      });
      Gmail.Users.Threads.get.mockImplementation((user, id) => ({
        id,
        messages: [makeMessage()]
      }));

      const threads = svc.getThreads('Home/Bills');

      expect(Gmail.Users.Threads.list).toHaveBeenCalledWith('me', {
        labelIds: ['Label_1'],
        maxResults: 100
      });
      expect(Gmail.Users.Threads.get).toHaveBeenCalledTimes(2);
      expect(threads).toHaveLength(2);
    });

    it('returns an empty array when there are no threads', () => {
      Gmail.Users.Threads.list.mockReturnValue({});
      const threads = svc.getThreads('Home/Bills');
      expect(threads).toEqual([]);
    });

    it('respects a custom maxResults value', () => {
      Gmail.Users.Threads.list.mockReturnValue({ threads: [] });
      svc.getThreads('Home/Bills', 50);
      expect(Gmail.Users.Threads.list).toHaveBeenCalledWith('me', {
        labelIds: ['Label_1'],
        maxResults: 50
      });
    });
  });

  // ── getFirstMessage ────────────────────────────────────────────────────────

  describe('getFirstMessage', () => {
    it('returns the first message from a thread', () => {
      const msg = makeMessage();
      const thread = { id: 'thread_1', messages: [msg, makeMessage({ id: 'msg_002' })] };
      expect(svc.getFirstMessage(thread)).toBe(msg);
    });
  });

  // ── getHeader ──────────────────────────────────────────────────────────────

  describe('getHeader', () => {
    it('returns the matching header value (case-insensitive)', () => {
      const msg = makeMessage();
      expect(svc.getHeader(msg, 'subject')).toBe('Your Bill is Ready');
      expect(svc.getHeader(msg, 'FROM')).toBe('billing@utility.com');
    });

    it('returns null when the header is not present', () => {
      const msg = makeMessage();
      expect(svc.getHeader(msg, 'X-Custom-Header')).toBeNull();
    });

    it('returns null when payload has no headers', () => {
      const msg = { payload: {} };
      expect(svc.getHeader(msg, 'Subject')).toBeNull();
    });
  });

  // ── getThreadSubject / getThreadFrom ───────────────────────────────────────

  describe('getThreadSubject', () => {
    it('returns the Subject of the first message', () => {
      const thread = { id: 't1', messages: [makeMessage()] };
      expect(svc.getThreadSubject(thread)).toBe('Your Bill is Ready');
    });

    it('returns empty string when Subject header is missing', () => {
      const msg = { id: 'm1', payload: { headers: [] } };
      const thread = { id: 't1', messages: [msg] };
      expect(svc.getThreadSubject(thread)).toBe('');
    });
  });

  describe('getThreadFrom', () => {
    it('returns the From address of the first message', () => {
      const thread = { id: 't1', messages: [makeMessage()] };
      expect(svc.getThreadFrom(thread)).toBe('billing@utility.com');
    });
  });

  // ── getPlainBody / getHtmlBody ─────────────────────────────────────────────

  describe('getPlainBody', () => {
    it('decodes base64url body for a simple text/plain message', () => {
      // Wire Utilities.base64DecodeWebSafe to return the real bytes
      Utilities.base64DecodeWebSafe.mockReturnValue(
        Array.from(Buffer.from('Due Date: 02/15/24\n$123.45\nACC123'))
      );
      Utilities.newBlob.mockReturnValue({
        getDataAsString: jest.fn(() => 'Due Date: 02/15/24\n$123.45\nACC123')
      });

      const body = svc.getPlainBody(makeMessage());
      expect(body).toBe('Due Date: 02/15/24\n$123.45\nACC123');
    });

    it('extracts text/plain from a multipart message', () => {
      Utilities.base64DecodeWebSafe.mockReturnValue(Array.from(Buffer.from('Plain text body')));
      Utilities.newBlob.mockReturnValue({
        getDataAsString: jest.fn(() => 'Plain text body')
      });

      const body = svc.getPlainBody(makeMultipartMessage());
      expect(body).toBe('Plain text body');
    });

    it('returns empty string when no matching part exists', () => {
      const msg = { payload: { mimeType: 'text/html', body: {} } };
      expect(svc.getPlainBody(msg)).toBe('');
    });
  });

  describe('getHtmlBody', () => {
    it('extracts text/html from a multipart message', () => {
      Utilities.base64DecodeWebSafe.mockReturnValue(Array.from(Buffer.from('<p>HTML body</p>')));
      Utilities.newBlob.mockReturnValue({
        getDataAsString: jest.fn(() => '<p>HTML body</p>')
      });

      const body = svc.getHtmlBody(makeMultipartMessage());
      expect(body).toBe('<p>HTML body</p>');
    });
  });

  // ── getMessageDate ─────────────────────────────────────────────────────────

  describe('getMessageDate', () => {
    it('parses date from the Date header', () => {
      const msg = makeMessage();
      const date = svc.getMessageDate(msg);
      expect(date).toBeInstanceOf(Date);
      expect(date.getFullYear()).toBe(2024);
    });

    it('falls back to internalDate when Date header is missing', () => {
      const msg = {
        internalDate: '1706784000000',
        payload: { headers: [] }
      };
      const date = svc.getMessageDate(msg);
      expect(date).toBeInstanceOf(Date);
      expect(date.getTime()).toBe(1706784000000);
    });
  });

  // ── getMessageId / getThreadId ─────────────────────────────────────────────

  describe('getMessageId', () => {
    it('returns message id', () => {
      expect(svc.getMessageId(makeMessage())).toBe('msg_001');
    });
  });

  describe('getThreadId', () => {
    it('returns thread id', () => {
      expect(svc.getThreadId({ id: 'thr_42', messages: [] })).toBe('thr_42');
    });
  });

  // ── addLabelToThread / moveThreadToInbox ───────────────────────────────────

  describe('addLabelToThread', () => {
    it('calls Gmail.Users.Threads.modify with the correct label', () => {
      svc.addLabelToThread('thread_A', 'Label_4');
      expect(Gmail.Users.Threads.modify).toHaveBeenCalledWith(
        { addLabelIds: ['Label_4'] }, 'me', 'thread_A'
      );
    });
  });

  describe('moveThreadToInbox', () => {
    it('adds the INBOX system label', () => {
      svc.moveThreadToInbox('thread_B');
      expect(Gmail.Users.Threads.modify).toHaveBeenCalledWith(
        { addLabelIds: ['INBOX'] }, 'me', 'thread_B'
      );
    });
  });

  // ── getAttachmentBody ──────────────────────────────────────────────────────

  describe('getAttachmentBody', () => {
    it('downloads attachment, converts to Doc, extracts text, removes temp file', () => {
      const pdfPart = {
        mimeType: 'application/pdf',
        filename: 'bill.pdf',
        body: { attachmentId: 'att_001' }
      };
      const msg = {
        id: 'msg_003',
        payload: { mimeType: 'multipart/mixed', parts: [pdfPart] }
      };
      Gmail.Users.Messages.Attachments.get.mockReturnValue({ data: b64('pdf content') });
      Utilities.base64DecodeWebSafe.mockReturnValue(Array.from(Buffer.from('pdf content')));
      Utilities.newBlob.mockReturnValue({
        getBytes: jest.fn(() => []),
        getContentType: jest.fn(() => 'application/pdf'),
        getName: jest.fn(() => 'bill.pdf')
      });

      const folder = {
        createFile: jest.fn(() => ({ getBlob: jest.fn(() => ({})) }))
      };

      const text = svc.getAttachmentBody(msg, 'msg_003', folder);
      expect(Gmail.Users.Messages.Attachments.get).toHaveBeenCalledWith('me', 'msg_003', 'att_001');
      expect(Drive.Files.create).toHaveBeenCalled();
      expect(Drive.Files.remove).toHaveBeenCalledWith('new_doc_id');
      expect(text).toBe('Due Date: 02/15/24\n$123.45\nACC123'); // from DocumentApp mock
    });

    it('throws when no PDF attachment part exists', () => {
      const msg = { id: 'msg_004', payload: { mimeType: 'text/plain', body: { data: '' } } };
      expect(() => svc.getAttachmentBody(msg, 'msg_004', {})).toThrow('No PDF attachment found');
    });
  });

  // ── sendEmail ──────────────────────────────────────────────────────────────

  describe('sendEmail', () => {
    it('builds a MIME message and calls Gmail.Users.Messages.send', () => {
      Utilities.base64EncodeWebSafe.mockReturnValue('encoded_mime');
      Utilities.getUuid.mockReturnValue('uuid-test');

      svc.sendEmail('to@example.com', 'Hello', 'Plain body', {
        cc: 'cc@example.com',
        name: 'Sender Name',
        htmlBody: '<p>HTML body</p>'
      });

      expect(Gmail.Users.Messages.send).toHaveBeenCalledWith(
        { raw: expect.any(String) }, 'me'
      );
    });

    it('includes inline image parts when inlineImages provided', () => {
      Utilities.base64EncodeWebSafe.mockReturnValue('encoded_with_images');
      Utilities.getUuid.mockReturnValue('uuid-img');
      Utilities.base64Encode.mockReturnValue('imgdata==');

      const blob = {
        getContentType: jest.fn(() => 'image/png'),
        getBytes: jest.fn(() => [1, 2, 3])
      };
      svc.sendEmail('to@example.com', 'Subject', 'Plain', {
        htmlBody: '<img src="cid:chart0">',
        inlineImages: { chart0: blob }
      });

      expect(Gmail.Users.Messages.send).toHaveBeenCalled();
    });

    it('sends plain-text-only when htmlBody is not provided', () => {
      Utilities.base64EncodeWebSafe.mockReturnValue('plain_only');
      Utilities.getUuid.mockReturnValue('uuid-plain');

      svc.sendEmail('to@example.com', 'Text Only', 'Just text');
      expect(Gmail.Users.Messages.send).toHaveBeenCalledWith({ raw: 'plain_only' }, 'me');
    });
  });

  // ── _extractBodyPart ───────────────────────────────────────────────────────

  describe('_extractBodyPart (private)', () => {
    it('returns empty string for null payload', () => {
      expect(svc._extractBodyPart(null, 'text/plain')).toBe('');
    });

    it('handles deeply nested parts', () => {
      Utilities.base64DecodeWebSafe.mockReturnValue(Array.from(Buffer.from('deep content')));
      Utilities.newBlob.mockReturnValue({ getDataAsString: jest.fn(() => 'deep content') });

      const payload = {
        mimeType: 'multipart/mixed',
        parts: [{
          mimeType: 'multipart/alternative',
          parts: [{
            mimeType: 'text/plain',
            body: { data: b64('deep content') }
          }]
        }]
      };
      expect(svc._extractBodyPart(payload, 'text/plain')).toBe('deep content');
    });
  });
});
