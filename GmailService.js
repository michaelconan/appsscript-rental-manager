/*
 * @module GmailService
 * @description Wrapper class for Gmail advanced service API operations.
 *              Replaces GmailApp standard library with direct Gmail REST API
 *              calls for improved performance and batch operation support.
 */

/**
 * Wrapper for the Gmail advanced service (Gmail REST API).
 * Provides label management, thread/message access, and email sending.
 */
class GmailService {
  constructor() {
    /** @private {Object} Cache of label name -> label ID */
    this._labelCache = {};
  }

  /**
   * Get Gmail label ID by display name, with in-memory caching.
   * @param {string} name - Gmail label name (e.g. 'Home/Bills')
   * @returns {string} Label ID
   * @throws {Error} If label is not found
   */
  getLabelId(name) {
    if (this._labelCache[name]) return this._labelCache[name];
    const response = Gmail.Users.Labels.list('me');
    const labels = (response && response.labels) || [];
    const label = labels.find(l => l.name === name);
    if (!label) throw new Error(`Gmail label not found: ${name}`);
    this._labelCache[name] = label.id;
    return label.id;
  }

  /**
   * Get all threads for a given label, returning full thread objects with messages.
   * @param {string} labelName - Gmail label display name
   * @param {number} [maxResults=100] - Max number of threads to return
   * @returns {Object[]} Array of full thread objects from Gmail API
   */
  getThreads(labelName, maxResults = 100) {
    const labelId = this.getLabelId(labelName);
    const response = Gmail.Users.Threads.list('me', { labelIds: [labelId], maxResults });
    const threads = (response && response.threads) || [];
    return threads.map(t => Gmail.Users.Threads.get('me', t.id, { format: 'full' }));
  }

  /**
   * Get the first message object from a thread.
   * @param {Object} thread - Full thread object from Gmail API
   * @returns {Object} First message object
   */
  getFirstMessage(thread) {
    return thread.messages[0];
  }

  /**
   * Get the subject line of the first message in a thread.
   * @param {Object} thread - Full thread object
   * @returns {string} Subject header value, or empty string
   */
  getThreadSubject(thread) {
    return this.getHeader(this.getFirstMessage(thread), 'Subject') || '';
  }

  /**
   * Get the sender address of the first message in a thread.
   * @param {Object} thread - Full thread object
   * @returns {string} From header value, or empty string
   */
  getThreadFrom(thread) {
    return this.getHeader(this.getFirstMessage(thread), 'From') || '';
  }

  /**
   * Extract a named header value from a Gmail message object.
   * @param {Object} message - Gmail API message object
   * @param {string} name - Header name (case-insensitive)
   * @returns {string|null} Header value, or null if not found
   */
  getHeader(message, name) {
    const headers = (message.payload && message.payload.headers) || [];
    const header = headers.find(h => h.name.toLowerCase() === name.toLowerCase());
    return header ? header.value : null;
  }

  /**
   * Extract the plain text body from a message, handling multipart payloads.
   * @param {Object} message - Gmail API message object
   * @returns {string} Plain text body content, or empty string
   */
  getPlainBody(message) {
    return this._extractBodyPart(message.payload, 'text/plain');
  }

  /**
   * Extract the HTML body from a message, handling multipart payloads.
   * @param {Object} message - Gmail API message object
   * @returns {string} HTML body content, or empty string
   */
  getHtmlBody(message) {
    return this._extractBodyPart(message.payload, 'text/html');
  }

  /**
   * Get the send date of a message.
   * @param {Object} message - Gmail API message object
   * @returns {Date} Message date parsed from Date header or internalDate
   */
  getMessageDate(message) {
    const dateStr = this.getHeader(message, 'Date');
    return dateStr ? new Date(dateStr) : new Date(parseInt(message.internalDate));
  }

  /**
   * Get the message ID string.
   * @param {Object} message - Gmail API message object
   * @returns {string} Message ID
   */
  getMessageId(message) {
    return message.id;
  }

  /**
   * Get the thread ID string.
   * @param {Object} thread - Gmail API thread object
   * @returns {string} Thread ID
   */
  getThreadId(thread) {
    return thread.id;
  }

  /**
   * Download a PDF attachment, convert to Google Doc, extract text, then clean up.
   * @param {Object} message - Gmail API message object (format: 'full')
   * @param {string} msgId - Message ID (required for Attachments.get)
   * @param {GoogleAppsScript.Drive.Folder} folder - Drive folder to store temp file
   * @returns {string} Extracted plain text from the PDF
   * @throws {Error} If no PDF attachment is found
   */
  getAttachmentBody(message, msgId, folder) {
    const attachmentPart = this._findAttachmentPart(message.payload);
    if (!attachmentPart) throw new Error('No PDF attachment found in message ' + msgId);

    const attachmentData = Gmail.Users.Messages.Attachments.get('me', msgId, attachmentPart.body.attachmentId);
    const blobData = Utilities.base64DecodeWebSafe(attachmentData.data);
    const blob = Utilities.newBlob(blobData, attachmentPart.mimeType, attachmentPart.filename);

    const billFile = folder.createFile(blob).getBlob();
    const newFile = Drive.Files.create(
      { title: attachmentPart.filename, mimeType: 'application/vnd.google-apps.document' },
      billFile
    );

    const text = DocumentApp.openById(newFile.id).getBody().getText();
    Drive.Files.remove(newFile.id);
    return text;
  }

  /**
   * Add a label to an entire thread.
   * @param {string} threadId - Gmail thread ID
   * @param {string} labelId - Gmail label ID to add
   */
  addLabelToThread(threadId, labelId) {
    Gmail.Users.Threads.modify({ addLabelIds: [labelId] }, 'me', threadId);
  }

  /**
   * Move a thread to the inbox by adding the INBOX system label.
   * @param {string} threadId - Gmail thread ID
   */
  moveThreadToInbox(threadId) {
    Gmail.Users.Threads.modify({ addLabelIds: ['INBOX'] }, 'me', threadId);
  }

  /**
   * Send an email via the Gmail advanced service using a raw MIME message.
   * Supports plain text, HTML body, inline images, cc, and sender name.
   * @param {string} to - Recipient email address
   * @param {string} subject - Email subject
   * @param {string} plainBody - Plain text version of body
   * @param {Object} [options={}] - Optional fields: cc, name, htmlBody, inlineImages
   */
  sendEmail(to, subject, plainBody, options = {}) {
    const fromEmail = Session.getEffectiveUser().getEmail();
    const fromHeader = options.name ? `${options.name} <${fromEmail}>` : fromEmail;
    const raw = this._buildMimeMessage(to, fromHeader, subject, plainBody, options);
    Gmail.Users.Messages.send({ raw }, 'me');
  }

  // ── Private helpers ──────────────────────────────────────────────────────────

  /**
   * Recursively search a message payload for a part matching the given MIME type
   * and decode its base64url-encoded body data.
   * @param {Object} payload - Gmail message payload or part
   * @param {string} mimeType - MIME type to find (e.g. 'text/plain')
   * @returns {string} Decoded text content, or empty string if not found
   * @private
   */
  _extractBodyPart(payload, mimeType) {
    if (!payload) return '';
    if (payload.mimeType === mimeType && payload.body && payload.body.data) {
      return Utilities.newBlob(Utilities.base64DecodeWebSafe(payload.body.data)).getDataAsString();
    }
    for (const part of (payload.parts || [])) {
      const text = this._extractBodyPart(part, mimeType);
      if (text) return text;
    }
    return '';
  }

  /**
   * Recursively search a message payload for a PDF attachment part.
   * @param {Object} payload - Gmail message payload or part
   * @returns {Object|null} The attachment part, or null if not found
   * @private
   */
  _findAttachmentPart(payload) {
    if (!payload) return null;
    if (payload.body && payload.body.attachmentId && payload.mimeType === 'application/pdf') {
      return payload;
    }
    for (const part of (payload.parts || [])) {
      const found = this._findAttachmentPart(part);
      if (found) return found;
    }
    return null;
  }

  /**
   * Build a base64url-encoded MIME message string for the Gmail API send endpoint.
   * Constructs multipart/related (inline images) > multipart/alternative (html+plain)
   * when inline images are present, otherwise multipart/alternative or plain text.
   * @param {string} to - Recipient
   * @param {string} from - Sender (with optional name)
   * @param {string} subject - Subject line
   * @param {string} plainBody - Plain text content
   * @param {Object} options - { cc, htmlBody, inlineImages }
   * @returns {string} base64url-encoded RFC 2822 message
   * @private
   */
  _buildMimeMessage(to, from, subject, plainBody, options) {
    const { cc, htmlBody, inlineImages } = options;
    const hasImages = inlineImages && Object.keys(inlineImages).length > 0;
    const relBoundary = 'rel_' + Utilities.getUuid();
    const altBoundary = 'alt_' + Utilities.getUuid();

    const lines = [
      `MIME-Version: 1.0`,
      `To: ${to}`,
      `From: ${from}`,
      `Subject: ${subject}`,
    ];
    if (cc) lines.push(`Cc: ${cc}`);

    if (hasImages && htmlBody) {
      lines.push(`Content-Type: multipart/related; boundary="${relBoundary}"`);
      lines.push('');
      lines.push(`--${relBoundary}`);
      lines.push(`Content-Type: multipart/alternative; boundary="${altBoundary}"`);
      lines.push('');
      lines.push(`--${altBoundary}`);
      lines.push('Content-Type: text/plain; charset=UTF-8');
      lines.push('');
      lines.push(plainBody);
      lines.push(`--${altBoundary}`);
      lines.push('Content-Type: text/html; charset=UTF-8');
      lines.push('');
      lines.push(htmlBody);
      lines.push(`--${altBoundary}--`);

      for (const [name, blob] of Object.entries(inlineImages)) {
        lines.push(`--${relBoundary}`);
        lines.push(`Content-Type: ${blob.getContentType()}`);
        lines.push('Content-Transfer-Encoding: base64');
        lines.push(`Content-ID: <${name}>`);
        lines.push(`Content-Disposition: inline; filename="${name}.png"`);
        lines.push('');
        lines.push(Utilities.base64Encode(blob.getBytes()));
      }
      lines.push(`--${relBoundary}--`);

    } else if (htmlBody) {
      lines.push(`Content-Type: multipart/alternative; boundary="${altBoundary}"`);
      lines.push('');
      lines.push(`--${altBoundary}`);
      lines.push('Content-Type: text/plain; charset=UTF-8');
      lines.push('');
      lines.push(plainBody);
      lines.push(`--${altBoundary}`);
      lines.push('Content-Type: text/html; charset=UTF-8');
      lines.push('');
      lines.push(htmlBody);
      lines.push(`--${altBoundary}--`);

    } else {
      lines.push('Content-Type: text/plain; charset=UTF-8');
      lines.push('');
      lines.push(plainBody);
    }

    return Utilities.base64EncodeWebSafe(lines.join('\r\n')).replace(/=+$/, '');
  }
}

// Node.js / Jest compatibility — ignored by Apps Script runtime
if (typeof module !== 'undefined') {
  module.exports = { GmailService };
}
