/**
 * RFQ Copilot taskpane: context-aware UI, preset RFQ data, email flows.
 */

const ENGINEERING_TEAM_EMAIL = 'engineering-team@Hexa729.onmicrosoft.com';
const CUSTOMER_1_EMAIL = 'customer-1@hexa729.onmicrosoft.com';
const OUTBOUND_SUBJECT_PREFIX = 'Technical Review Required - RFQ #41260018 (NRL - 2 FBG Arrays)';
const FIND_MESSAGE_DELAY_MS = 2500;
const FIND_MESSAGE_RETRIES = 3;
const FIND_MESSAGE_RETRY_DELAY_MS = 1500;
const CUSTOMER_1_FIND_DELAY_MS = 4000;
const CUSTOMER_1_FIND_RETRIES = 4;
const CUSTOMER_1_FIND_RETRY_DELAY_MS = 2000;
const STORAGE_KEY = 'addin_original_message_map';
const STORAGE_MAX_ENTRIES = 50;

var currentOriginalMessageRestId = null;

function makeUniqueSubject() {
  return OUTBOUND_SUBJECT_PREFIX + ' – ' + new Date().toISOString();
}

function getOriginalMessageMap() {
  try {
    var raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch (e) {
    return {};
  }
}

function setOriginalMessageMap(map) {
  var keys = Object.keys(map);
  if (keys.length > STORAGE_MAX_ENTRIES) {
    var sorted = keys.sort();
    for (var i = 0; i < sorted.length - STORAGE_MAX_ENTRIES; i++) {
      delete map[sorted[i]];
    }
  }
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(map));
  } catch (e) {}
}

function saveOriginalForSubject(subject, restId) {
  var map = getOriginalMessageMap();
  map[subject] = restId;
  setOriginalMessageMap(map);
}

Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    initializeApp().catch(function (err) {
      console.error('initializeApp error:', err);
      showStatus('Initialization failed: ' + (err && err.message), true);
    });
  }
});

async function initializeApp() {
  setupEventListeners();
  await initializeAuth();
  updateAuthUI();
  updateSendButtonState();
  renderInitialRfqFromPreset();
  renderEngineeringAnswersFromPreset();
  detectEmailContext();
  if (Office.context.mailbox && Office.context.mailbox.addHandlerAsync) {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () {
      detectEmailContext();
    }, function (err) {
      if (err) console.warn('ItemChanged handler not registered:', err);
    });
  }
  setInterval(function () {
    detectEmailContext();
  }, 2000);
}

async function initializeAuth() {
  try {
    var initialized = await AuthService.initialize();
    if (!initialized) console.error('Auth not initialized');
  } catch (error) {
    console.error('Auth initialization failed:', error);
    throw error;
  }
}

function updateAuthUI() {
  var signInBtn = document.getElementById('sign-in-btn');
  var userName = document.getElementById('user-name');
  var accountDivider = document.getElementById('account-divider');
  var signOutBtn = document.getElementById('sign-out-btn');

  if (AuthService.isSignedIn()) {
    var user = AuthService.getUser();
    if (signInBtn) signInBtn.classList.add('hidden');
    if (userName) {
      userName.classList.remove('hidden');
      userName.textContent = user ? (user.name || user.email) : '';
      userName.setAttribute('title', user ? user.email : '');
    }
    if (accountDivider) accountDivider.classList.remove('hidden');
    if (signOutBtn) signOutBtn.classList.remove('hidden');
  } else {
    if (signInBtn) signInBtn.classList.remove('hidden');
    if (userName) userName.classList.add('hidden');
    if (accountDivider) accountDivider.classList.add('hidden');
    if (signOutBtn) signOutBtn.classList.add('hidden');
  }
}

function updateSendButtonState() {
  var btn = document.getElementById('send-rfq-btn');
  if (btn) btn.disabled = !AuthService.isSignedIn();
  var replyBtn = document.getElementById('reply-to-original-btn');
  if (replyBtn) replyBtn.disabled = !AuthService.isSignedIn();
  var finalReplyBtn = document.getElementById('send-final-reply-btn');
  if (finalReplyBtn) finalReplyBtn.disabled = !AuthService.isSignedIn();
  var directReplyBtn = document.getElementById('reply-direct-to-customer-btn');
  if (directReplyBtn) directReplyBtn.disabled = !AuthService.isSignedIn();
}

function renderInitialRfqFromPreset() {
  var data = typeof RFQ_DATA !== 'undefined' ? RFQ_DATA : null;
  if (!data) return;

  var grid = document.getElementById('specs-grid');
  if (grid) {
    var s = data.technicalSpecs;
    var specs = [
      { label: 'Quantity', value: s.quantity },
      { label: 'Fiber type', value: s.fiberType },
      { label: 'FBGs', value: s.numberOfFBGs },
      { label: 'Total length', value: s.totalFiberLength },
      { label: 'Wavelength', value: s.wavelengthNm + ' nm' },
      { label: 'Spacing', value: s.fbgSpacingMm + ' mm' },
      { label: 'Connector', value: s.connectorType },
      { label: 'Coating', value: s.coatingMaterial },
    ];
    grid.innerHTML = specs.map(function (sp) {
      return '<div class="spec-item"><span class="spec-label">' + escapeHtml(sp.label) + '</span><span class="spec-value">' + escapeHtml(String(sp.value)) + '</span></div>';
    }).join('');
  }

  var accordion = document.getElementById('customer-questions-accordion');
  if (accordion && data.customerQuestions && data.customerQuestions.length) {
    accordion.innerHTML = data.customerQuestions.map(function (q, idx) {
      var id = 'accordion-' + (q.id || idx);
      return (
        '<div class="accordion-item" data-accordion-id="' + id + '">' +
          '<button type="button" class="accordion-header" aria-expanded="false">' +
            '<span>' + escapeHtml((idx + 1) + '. ' + q.question) + '</span>' +
            '<span class="chevron" aria-hidden="true">▼</span>' +
          '</button>' +
          '<div class="accordion-body">' +
            '<div class="accordion-body-inner">' + escapeHtml(q.answer) + '</div>' +
          '</div>' +
        '</div>'
      );
    }).join('');
    accordion.querySelectorAll('.accordion-header').forEach(function (btn) {
      btn.addEventListener('click', function () {
        var item = btn.closest('.accordion-item');
        if (item) item.classList.toggle('open');
      });
    });
  }

  var missingList = document.getElementById('missing-details-list');
  if (missingList && data.missingDetails && data.missingDetails.length) {
    missingList.innerHTML = data.missingDetails.map(function (m) {
      var badgeClass = m.importance === 'critical' ? 'badge-critical' : 'badge-high';
      return '<li><span class="badge ' + badgeClass + '">' + escapeHtml(m.importance || '') + '</span> ' + escapeHtml(m.field) + '</li>';
    }).join('');
  }

  var askList = document.getElementById('questions-to-ask-list');
  if (askList && data.questionsToAsk && data.questionsToAsk.length) {
    askList.innerHTML = data.questionsToAsk.map(function (q) {
      return '<li>' + escapeHtml(q) + '</li>';
    }).join('');
  }
}

function renderEngineeringAnswersFromPreset() {
  var templates = typeof EMAIL_TEMPLATES !== 'undefined' ? EMAIL_TEMPLATES : null;
  if (!templates || !templates.engineeringReply || !templates.engineeringReply.body) return;

  var container = document.getElementById('engineering-answers-container');
  if (!container) return;

  var body = templates.engineeringReply.body;
  var lines = body.split(/\n/).filter(function (l) { return l.trim(); });
  var cards = [];
  var current = [];
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    var numMatch = line.match(/^(\d+)\.\s/);
    if (numMatch && current.length) {
      cards.push(current.join('\n'));
      current = [line];
    } else {
      current.push(line);
    }
  }
  if (current.length) cards.push(current.join('\n'));

  container.innerHTML = cards.map(function (text, idx) {
    return '<div class="card"><div class="answer-number">Item ' + (idx + 1) + '</div><div class="answer-text">' + escapeHtml(text) + '</div></div>';
  }).join('');
}

function escapeHtml(s) {
  if (s == null) return '';
  var div = document.createElement('div');
  div.textContent = s;
  return div.innerHTML;
}

function showContext(contextId) {
  var blocks = ['context-neutral', 'context-initial-rfq', 'context-engineering-reply', 'context-customer-reply-details'];
  blocks.forEach(function (id) {
    var el = document.getElementById(id);
    if (el) el.classList.toggle('visible', id === contextId);
  });
}

function detectEmailContext() {
  var statusEl = document.getElementById('status-message');
  currentOriginalMessageRestId = null;

  if (!Office.context.mailbox || !Office.context.mailbox.item) {
    showContext('context-neutral');
    if (statusEl) {
      statusEl.textContent = 'Open an RFQ from customer-1 or a reply from Engineering to use the copilot.';
      statusEl.classList.remove('error', 'success');
    }
    return;
  }

  var item = Office.context.mailbox.item;
  var subject = (item.subject || '').trim();
  var fromAddress = '';
  if (item.from) {
    var raw = item.from.emailAddress;
    if (typeof raw === 'string') fromAddress = raw.toLowerCase();
    else if (raw && typeof raw.address === 'string') fromAddress = raw.address.toLowerCase();
  }
  var normalizedSubject = (item.normalizedSubject != null ? item.normalizedSubject : subject.replace(/^Re:\s*/i, '')).trim();
  var subjectLower = subject.toLowerCase();

  var isFromEngineering = fromAddress === ENGINEERING_TEAM_EMAIL.toLowerCase();
  var isFromCustomer1 = fromAddress === CUSTOMER_1_EMAIL.toLowerCase();
  var isReplyFromEngineering = isFromEngineering && normalizedSubject.indexOf(OUTBOUND_SUBJECT_PREFIX) === 0;
  var isReplyFromCustomer1 = isFromCustomer1 && subjectLower.indexOf('re:') === 0;

  if (isReplyFromEngineering) {
    var map = getOriginalMessageMap();
    var originalRestId = map[normalizedSubject];
    if (originalRestId) currentOriginalMessageRestId = originalRestId;
    showContext('context-engineering-reply');
    if (statusEl) { statusEl.textContent = ''; statusEl.classList.remove('error', 'success'); }
    if (!originalRestId && statusEl) {
      statusEl.textContent = 'Original RFQ message not found for this thread.';
      statusEl.classList.add('error');
    }
    return;
  }

  if (isReplyFromCustomer1) {
    showContext('context-customer-reply-details');
    if (statusEl) { statusEl.textContent = ''; statusEl.classList.remove('error', 'success'); }
    return;
  }

  if (isFromCustomer1) {
    showContext('context-initial-rfq');
    if (statusEl) { statusEl.textContent = ''; statusEl.classList.remove('error', 'success'); }
    return;
  }

  showContext('context-neutral');
  if (statusEl) {
    statusEl.textContent = 'Open an RFQ from customer-1 or a reply from Engineering to use the copilot.';
    statusEl.classList.remove('error', 'success');
  }
}

function showStatus(message, isError) {
  var el = document.getElementById('status-message');
  if (!el) return;
  el.textContent = message || '';
  el.classList.remove('error', 'success');
  if (isError) el.classList.add('error');
  else if (message) el.classList.add('success');
}

function setupEventListeners() {
  document.getElementById('sign-in-btn')?.addEventListener('click', handleSignIn);
  document.getElementById('sign-out-btn')?.addEventListener('click', handleSignOut);
  document.getElementById('send-rfq-btn')?.addEventListener('click', handleSendRfq);
  document.getElementById('reply-to-original-btn')?.addEventListener('click', handleReplyToOriginal);
  document.getElementById('send-final-reply-btn')?.addEventListener('click', handleSendFinalReplyToCustomer);
  document.getElementById('reply-direct-to-customer-btn')?.addEventListener('click', handleReplyDirectToCustomer);
}

async function handleSignIn() {
  try {
    showStatus('Signing in...', false);
    await AuthService.signIn();
    updateAuthUI();
    updateSendButtonState();
    showStatus('Signed in successfully.', false);
  } catch (error) {
    console.error('Sign in error:', error);
    showStatus('Sign in failed: ' + (error && error.message), true);
  }
}

async function handleSignOut() {
  try {
    await AuthService.signOut();
    updateAuthUI();
    updateSendButtonState();
    showStatus('', false);
  } catch (error) {
    console.error('Sign out error:', error);
    showStatus('Sign out failed: ' + (error && error.message), true);
  }
}

function getEngineeringReviewBody() {
  return (typeof EMAIL_TEMPLATES !== 'undefined' && EMAIL_TEMPLATES.engineeringReview && EMAIL_TEMPLATES.engineeringReview.body)
    ? EMAIL_TEMPLATES.engineeringReview.body
    : 'Key concerns requiring engineering input.';
}

function getEngineeringReplyComment() {
  return (typeof EMAIL_TEMPLATES !== 'undefined' && EMAIL_TEMPLATES.engineeringReply && EMAIL_TEMPLATES.engineeringReply.body)
    ? EMAIL_TEMPLATES.engineeringReply.body
    : 'Engineering assessment completed.';
}

function getCustomerClarificationComment() {
  return (typeof EMAIL_TEMPLATES !== 'undefined' && EMAIL_TEMPLATES.customerClarification && EMAIL_TEMPLATES.customerClarification.body)
    ? EMAIL_TEMPLATES.customerClarification.body
    : 'Thank you for your RFQ. We will respond with clarifications shortly.';
}

function getCustomerReplyWithDetailsComment() {
  return (typeof EMAIL_TEMPLATES !== 'undefined' && EMAIL_TEMPLATES.customerReplyWithDetails && EMAIL_TEMPLATES.customerReplyWithDetails.body)
    ? EMAIL_TEMPLATES.customerReplyWithDetails.body
    : 'Thank you for the clarifications. We have received your details.';
}

function getFinalQuoteComment() {
  return (typeof EMAIL_TEMPLATES !== 'undefined' && EMAIL_TEMPLATES.finalQuoteToCustomer && EMAIL_TEMPLATES.finalQuoteToCustomer.body)
    ? EMAIL_TEMPLATES.finalQuoteToCustomer.body
    : 'Please find attached our quote. Thank you.';
}

async function handleSendRfq() {
  if (!AuthService.isSignedIn()) {
    showStatus('Please sign in.', true);
    return;
  }

  var user = AuthService.getUser();
  var userEmail = user ? user.email : '';

  var outboundSubject = makeUniqueSubject();
  if (Office.context.mailbox && Office.context.mailbox.item) {
    try {
      var itemId = Office.context.mailbox.item.itemId;
      if (itemId && Office.context.mailbox.convertToRestId) {
        var restId = Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);
        if (restId) saveOriginalForSubject(outboundSubject, restId);
      }
    } catch (e) {
      console.warn('Could not store original message id:', e);
    }
  }

  try {
    showStatus('Sending to Engineering...', false);
    var sendRfqBtn = document.getElementById('send-rfq-btn');
    if (sendRfqBtn) sendRfqBtn.disabled = true;

    var message = {
      subject: outboundSubject,
      body: { contentType: 'Text', content: getEngineeringReviewBody() },
      toRecipients: [{ emailAddress: { address: ENGINEERING_TEAM_EMAIL, name: 'Engineering Team' } }],
    };
    await AuthService.graphRequest('/me/sendMail', {
      method: 'POST',
      body: JSON.stringify({ message: message, saveToSentItems: true }),
    });

    var foundMessage = null;
    for (var attempt = 0; attempt < FIND_MESSAGE_RETRIES; attempt++) {
      await sleep(attempt === 0 ? FIND_MESSAGE_DELAY_MS : FIND_MESSAGE_RETRY_DELAY_MS);
      var inboxUrl = '/users/' + encodeURIComponent(ENGINEERING_TEAM_EMAIL) + '/mailFolders/inbox/messages?' +
        '$orderby=receivedDateTime desc&$top=20&$select=id,subject,from,receivedDateTime';
      var result = await AuthService.graphRequest(inboxUrl);
      var messages = (result && result.value) || [];
      for (var i = 0; i < messages.length; i++) {
        var msg = messages[i];
        var fromAddr = msg.from && msg.from.emailAddress && msg.from.emailAddress.address;
        if (msg.subject === outboundSubject && fromAddr === userEmail) {
          foundMessage = msg;
          break;
        }
      }
      if (foundMessage) break;
    }

    if (!foundMessage) {
      showStatus('Could not find the sent message in engineering-team inbox. Ensure you have Full Access to that mailbox.', true);
      return;
    }

    var replyUrl = '/users/' + encodeURIComponent(ENGINEERING_TEAM_EMAIL) + '/messages/' + encodeURIComponent(foundMessage.id) + '/reply';
    await AuthService.graphRequest(replyUrl, {
      method: 'POST',
      body: JSON.stringify({ comment: getEngineeringReplyComment() }),
    });

    showStatus('Sent to Engineering and reply received in the same thread.', false);
  } catch (error) {
    console.error('Send RFQ error:', error);
    var msg = error && error.message ? error.message : String(error);
    var errLower = msg.toLowerCase();
    if (errLower.indexOf('not found') !== -1 || errLower.indexOf('404') !== -1 || errLower.indexOf('default folder') !== -1) {
      msg = 'Email was sent to engineering-team. To enable the auto-reply, your admin must grant you Full Access (and Send As) to engineering-team@Hexa729.onmicrosoft.com.';
    } else if (msg.indexOf('403') !== -1 || msg.indexOf('Access') !== -1) {
      msg = 'Access denied. In Azure, add delegated permission Mail.Send.Shared (Microsoft Graph), then sign out and sign in again. Ensure Full Access and Send As (or Send on behalf) for engineering-team@Hexa729.onmicrosoft.com in Exchange.';
    }
    showStatus(msg, true);
  } finally {
    var btn = document.getElementById('send-rfq-btn');
    if (btn) btn.disabled = !AuthService.isSignedIn();
  }
}

async function handleReplyDirectToCustomer() {
  if (!AuthService.isSignedIn()) {
    showStatus('Please sign in.', true);
    return;
  }
  if (!Office.context.mailbox || !Office.context.mailbox.item) {
    showStatus('No email selected.', true);
    return;
  }

  var user = AuthService.getUser();
  var userEmail = user ? user.email : '';
  var item = Office.context.mailbox.item;
  var itemId = item.itemId;
  if (!itemId || !Office.context.mailbox.convertToRestId) {
    showStatus('Could not get message id.', true);
    return;
  }

  var originalRestId;
  try {
    originalRestId = Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);
  } catch (e) {
    showStatus('Could not get message id.', true);
    return;
  }
  if (!originalRestId) {
    showStatus('Could not get message id.', true);
    return;
  }

  var directBtn = document.getElementById('reply-direct-to-customer-btn');
  if (directBtn) directBtn.disabled = true;
  var replySent = false;

  try {
    showStatus('Sending reply...', false);

    var originalSubject = (item.subject || '').trim();
    if (!originalSubject) {
      var originalMsg = await AuthService.graphRequest('/me/messages/' + encodeURIComponent(originalRestId) + '?$select=subject');
      originalSubject = (originalMsg && originalMsg.subject) ? originalMsg.subject.trim() : '';
    }

    await AuthService.graphRequest('/me/messages/' + encodeURIComponent(originalRestId) + '/reply', {
      method: 'POST',
      body: JSON.stringify({ comment: getCustomerClarificationComment() }),
    });
    replySent = true;

    var replySubject = 'Re: ' + originalSubject;
    var replySubjectLower = replySubject.trim().toLowerCase();
    var originalSubjectLower = originalSubject.trim().toLowerCase();
    var userEmailLower = (userEmail || '').toLowerCase();
    var foundInCustomer1 = null;
    for (var attempt = 0; attempt < CUSTOMER_1_FIND_RETRIES; attempt++) {
      await sleep(attempt === 0 ? CUSTOMER_1_FIND_DELAY_MS : CUSTOMER_1_FIND_RETRY_DELAY_MS);
      var inboxUrl = '/users/' + encodeURIComponent(CUSTOMER_1_EMAIL) + '/mailFolders/inbox/messages?' +
        '$orderby=receivedDateTime desc&$top=20&$select=id,subject,from,receivedDateTime';
      var result = await AuthService.graphRequest(inboxUrl);
      var messages = (result && result.value) || [];
      for (var i = 0; i < messages.length; i++) {
        var msg = messages[i];
        var fromAddr = msg.from && msg.from.emailAddress && msg.from.emailAddress.address;
        var fromLower = (fromAddr || '').toLowerCase();
        var msgSubjectNorm = (msg.subject || '').trim().toLowerCase();
        var isSubjectMatch = msgSubjectNorm === replySubjectLower ||
          (msgSubjectNorm.indexOf('re:') === 0 && msgSubjectNorm.replace(/^re:\s*/i, '').trim() === originalSubjectLower);
        if (isSubjectMatch && fromLower === userEmailLower) {
          foundInCustomer1 = msg;
          break;
        }
      }
      if (foundInCustomer1) break;
    }

    if (!foundInCustomer1) {
      showStatus('Reply sent. Could not find your reply in customer-1 inbox. Ensure Full Access to customer-1@hexa729.onmicrosoft.com.', true);
      return;
    }

    var customer1ReplyUrl = '/users/' + encodeURIComponent(CUSTOMER_1_EMAIL) + '/messages/' + encodeURIComponent(foundInCustomer1.id) + '/reply';
    await AuthService.graphRequest(customer1ReplyUrl, {
      method: 'POST',
      body: JSON.stringify({ comment: getCustomerReplyWithDetailsComment() }),
    });

    showStatus('Reply sent and automated reply from customer received. Open that reply to send the final quote.', false);
  } catch (error) {
    console.error('Reply direct to customer error:', error);
    var msg = error && error.message ? error.message : String(error);
    if (replySent) {
      var errLower = msg.toLowerCase();
      if (errLower.indexOf('not found') !== -1 || errLower.indexOf('404') !== -1 || errLower.indexOf('default folder') !== -1) {
        showStatus('Reply sent. Automated reply from customer-1 failed: ensure Full Access and Send As (or Send on behalf) for customer-1@hexa729.onmicrosoft.com.', true);
      } else if (msg.indexOf('403') !== -1 || msg.indexOf('Access') !== -1) {
        showStatus('Reply sent. Automated reply from customer-1 failed: ensure Send As (or Send on behalf) for customer-1@hexa729.onmicrosoft.com, then sign in again.', true);
      } else {
        showStatus('Reply sent. Automated reply from customer-1 failed: ' + msg, true);
      }
    } else {
      showStatus('Reply failed: ' + msg, true);
    }
  } finally {
    var btn = document.getElementById('reply-direct-to-customer-btn');
    if (btn) btn.disabled = !AuthService.isSignedIn();
  }
}

async function handleReplyToOriginal() {
  if (!AuthService.isSignedIn()) {
    showStatus('Please sign in.', true);
    return;
  }

  var user = AuthService.getUser();
  var userEmail = user ? user.email : '';

  var originalRestId = currentOriginalMessageRestId;
  if (!originalRestId && Office.context.mailbox && Office.context.mailbox.item) {
    var item = Office.context.mailbox.item;
    var subject = (item.subject || '').trim();
    var normalizedSubject = subject.replace(/^Re:\s*/i, '').trim();
    var map = getOriginalMessageMap();
    originalRestId = map[normalizedSubject];
  }

  if (!originalRestId) {
    showStatus('Original message not found. Cannot send clarifications.', true);
    return;
  }

  var replyBtn = document.getElementById('reply-to-original-btn');
  if (replyBtn) replyBtn.disabled = true;
  var replyToOriginalSent = false;

  try {
    showStatus('Sending clarifications to customer...', false);

    var originalMsg = await AuthService.graphRequest('/me/messages/' + encodeURIComponent(originalRestId) + '?$select=subject');
    var originalSubject = (originalMsg && originalMsg.subject) ? originalMsg.subject.trim() : '';

    await AuthService.graphRequest('/me/messages/' + encodeURIComponent(originalRestId) + '/reply', {
      method: 'POST',
      body: JSON.stringify({ comment: getCustomerClarificationComment() }),
    });
    replyToOriginalSent = true;

    var replySubject = 'Re: ' + originalSubject;
    var replySubjectLower = replySubject.trim().toLowerCase();
    var originalSubjectLower = originalSubject.trim().toLowerCase();
    var userEmailLower = (userEmail || '').toLowerCase();
    var foundInCustomer1 = null;
    for (var attempt = 0; attempt < CUSTOMER_1_FIND_RETRIES; attempt++) {
      await sleep(attempt === 0 ? CUSTOMER_1_FIND_DELAY_MS : CUSTOMER_1_FIND_RETRY_DELAY_MS);
      var inboxUrl = '/users/' + encodeURIComponent(CUSTOMER_1_EMAIL) + '/mailFolders/inbox/messages?' +
        '$orderby=receivedDateTime desc&$top=20&$select=id,subject,from,receivedDateTime';
      var result = await AuthService.graphRequest(inboxUrl);
      var messages = (result && result.value) || [];
      for (var i = 0; i < messages.length; i++) {
        var msg = messages[i];
        var fromAddr = msg.from && msg.from.emailAddress && msg.from.emailAddress.address;
        var fromLower = (fromAddr || '').toLowerCase();
        var msgSubjectNorm = (msg.subject || '').trim().toLowerCase();
        var isSubjectMatch = msgSubjectNorm === replySubjectLower ||
          (msgSubjectNorm.indexOf('re:') === 0 && msgSubjectNorm.replace(/^re:\s*/i, '').trim() === originalSubjectLower);
        if (isSubjectMatch && fromLower === userEmailLower) {
          foundInCustomer1 = msg;
          break;
        }
      }
      if (foundInCustomer1) break;
    }

    if (!foundInCustomer1) {
      showStatus('Clarifications sent. Could not find your reply in customer-1 inbox. Ensure Full Access to customer-1@hexa729.onmicrosoft.com.', true);
      return;
    }

    var customer1ReplyUrl = '/users/' + encodeURIComponent(CUSTOMER_1_EMAIL) + '/messages/' + encodeURIComponent(foundInCustomer1.id) + '/reply';
    await AuthService.graphRequest(customer1ReplyUrl, {
      method: 'POST',
      body: JSON.stringify({ comment: getCustomerReplyWithDetailsComment() }),
    });

    showStatus('Clarifications sent and automated reply from customer received. Open that reply to send the final quote.', false);
  } catch (error) {
    console.error('Reply to original error:', error);
    var msg = error && error.message ? error.message : String(error);
    var errLower = msg.toLowerCase();
    if (replyToOriginalSent) {
      if (errLower.indexOf('not found') !== -1 || errLower.indexOf('404') !== -1 || errLower.indexOf('default folder') !== -1) {
        showStatus('Clarifications sent. Automated reply from customer-1 failed: ensure Full Access and Send As (or Send on behalf) for customer-1@hexa729.onmicrosoft.com.', true);
      } else if (msg.indexOf('403') !== -1 || msg.indexOf('Access') !== -1) {
        showStatus('Clarifications sent. Automated reply from customer-1 failed: ensure Send As (or Send on behalf) for customer-1@hexa729.onmicrosoft.com, then sign in again.', true);
      } else {
        showStatus('Clarifications sent. Automated reply from customer-1 failed: ' + msg, true);
      }
    } else {
      showStatus('Send failed: ' + msg, true);
    }
  } finally {
    var btn = document.getElementById('reply-to-original-btn');
    if (btn) btn.disabled = !AuthService.isSignedIn();
  }
}

async function handleSendFinalReplyToCustomer() {
  if (!AuthService.isSignedIn()) {
    showStatus('Please sign in.', true);
    return;
  }
  if (!Office.context.mailbox || !Office.context.mailbox.item) {
    showStatus('No email selected.', true);
    return;
  }

  var itemId = Office.context.mailbox.item.itemId;
  if (!itemId || !Office.context.mailbox.convertToRestId) {
    showStatus('Could not get message id.', true);
    return;
  }

  try {
    var restId = Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);
    if (!restId) {
      showStatus('Could not get message id.', true);
      return;
    }

    showStatus('Sending final quote...', false);
    var btn = document.getElementById('send-final-reply-btn');
    if (btn) btn.disabled = true;

    await AuthService.graphRequest('/me/messages/' + encodeURIComponent(restId) + '/reply', {
      method: 'POST',
      body: JSON.stringify({ comment: getFinalQuoteComment() }),
    });

    showStatus('Final quote sent to customer.', false);
  } catch (error) {
    console.error('Send final reply error:', error);
    showStatus('Final quote failed: ' + (error && error.message ? error.message : String(error)), true);
  } finally {
    var finalBtn = document.getElementById('send-final-reply-btn');
    if (finalBtn) finalBtn.disabled = !AuthService.isSignedIn();
  }
}

function sleep(ms) {
  return new Promise(function (resolve) {
    setTimeout(resolve, ms);
  });
}
