/**
 * Taskpane: Office.onReady → initializeApp (auth first), then Send button and email flow.
 */

const ENGINEERING_TEAM_EMAIL = 'engineering-team@Hexa729.onmicrosoft.com';
const CUSTOMER_1_EMAIL = 'customer-1@hexa729.onmicrosoft.com';
const OUTBOUND_SUBJECT_PREFIX = 'Add-in RFQ notification';
const OUTBOUND_BODY = 'This is an automated notification from the Outlook add-in.';
const REPLY_COMMENT = 'We have received your request. Reference: engineering-team.';
const REPLY_TO_ORIGINAL_COMMENT = 'Replied from add-in after receiving engineering response.';
const CUSTOMER_1_AUTO_REPLY_COMMENT = 'Thank you for your reply. We have received the update. - customer-1';
const FINAL_REPLY_TO_CUSTOMER_COMMENT = 'This concludes the thread. Thank you. - Add-in';
const REPLY_DIRECT_TO_CUSTOMER_COMMENT = 'Thank you for your RFQ. We have received it and will respond. - Add-in';
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
    const initialized = await AuthService.initialize();
    if (!initialized) {
      console.error('Auth not initialized');
    }
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

function detectEmailContext() {
  var notifyBlock = document.getElementById('context-notify-engineering');
  var replyBlock = document.getElementById('context-reply-to-original');
  var replyBtn = document.getElementById('reply-to-original-btn');
  var finalReplyBlock = document.getElementById('context-final-reply-to-customer');
  var finalReplyBtn = document.getElementById('send-final-reply-btn');
  var statusEl = document.getElementById('status-message');
  currentOriginalMessageRestId = null;

  if (!Office.context.mailbox || !Office.context.mailbox.item) {
    if (notifyBlock) notifyBlock.classList.add('hidden');
    if (replyBlock) replyBlock.classList.add('hidden');
    if (finalReplyBlock) finalReplyBlock.classList.add('hidden');
    if (statusEl) {
      statusEl.textContent = 'Open an email from customer-1 or the reply from Engineering to use actions.';
      statusEl.classList.remove('error', 'success');
    }
    return;
  }

  var item = Office.context.mailbox.item;
  var subject = (item.subject || '').trim();
  var fromAddress = '';
  if (item.from) {
    var raw = item.from.emailAddress;
    if (typeof raw === 'string') {
      fromAddress = raw.toLowerCase();
    } else if (raw && typeof raw.address === 'string') {
      fromAddress = raw.address.toLowerCase();
    }
  }
  var normalizedSubject = (item.normalizedSubject != null ? item.normalizedSubject : subject.replace(/^Re:\s*/i, '')).trim();
  var subjectLower = subject.toLowerCase();

  var isReplyFromEngineering =
    fromAddress === ENGINEERING_TEAM_EMAIL.toLowerCase() &&
    normalizedSubject.indexOf(OUTBOUND_SUBJECT_PREFIX) === 0;

  var isReplyFromCustomer1 =
    fromAddress === CUSTOMER_1_EMAIL.toLowerCase() &&
    subjectLower.indexOf('re:') === 0;

  if (isReplyFromEngineering) {
    var map = getOriginalMessageMap();
    var originalRestId = map[normalizedSubject];
    if (originalRestId) {
      currentOriginalMessageRestId = originalRestId;
      if (notifyBlock) notifyBlock.classList.add('hidden');
      if (replyBlock) replyBlock.classList.remove('hidden');
      if (finalReplyBlock) finalReplyBlock.classList.add('hidden');
      if (replyBtn) replyBtn.disabled = !AuthService.isSignedIn();
      if (statusEl) {
        statusEl.textContent = '';
        statusEl.classList.remove('error', 'success');
      }
    } else {
      if (notifyBlock) notifyBlock.classList.add('hidden');
      if (replyBlock) replyBlock.classList.remove('hidden');
      if (finalReplyBlock) finalReplyBlock.classList.add('hidden');
      if (replyBtn) replyBtn.disabled = true;
      if (statusEl) {
        statusEl.textContent = 'Original message not found.';
        statusEl.classList.add('error');
      }
    }
  } else if (isReplyFromCustomer1) {
    if (notifyBlock) notifyBlock.classList.add('hidden');
    if (replyBlock) replyBlock.classList.add('hidden');
    if (finalReplyBlock) finalReplyBlock.classList.remove('hidden');
    if (finalReplyBtn) finalReplyBtn.disabled = !AuthService.isSignedIn();
    if (statusEl) {
      statusEl.textContent = '';
      statusEl.classList.remove('error', 'success');
    }
  } else {
    var isFromCustomer1 = fromAddress === CUSTOMER_1_EMAIL.toLowerCase();
    if (isFromCustomer1) {
      if (notifyBlock) notifyBlock.classList.remove('hidden');
      if (replyBlock) replyBlock.classList.add('hidden');
      if (finalReplyBlock) finalReplyBlock.classList.add('hidden');
      if (statusEl && statusEl.textContent !== '') {
        statusEl.textContent = '';
        statusEl.classList.remove('error', 'success');
      }
    } else {
      if (notifyBlock) notifyBlock.classList.add('hidden');
      if (replyBlock) replyBlock.classList.add('hidden');
      if (finalReplyBlock) finalReplyBlock.classList.add('hidden');
      if (statusEl) {
        statusEl.textContent = 'Open an RFQ from customer-1 or the reply from Engineering to use actions.';
        statusEl.classList.remove('error', 'success');
      }
    }
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
    showStatus('Sending...', false);
    var sendRfqBtn = document.getElementById('send-rfq-btn');
    if (sendRfqBtn) sendRfqBtn.disabled = true;

    // Step 1: Send a new email from current user to engineering-team (unique subject per click)
    var message = {
      subject: outboundSubject,
      body: {
        contentType: 'Text',
        content: OUTBOUND_BODY,
      },
      toRecipients: [
        {
          emailAddress: {
            address: ENGINEERING_TEAM_EMAIL,
            name: 'Engineering Team',
          },
        },
      ],
    };
    await AuthService.graphRequest('/me/sendMail', {
      method: 'POST',
      body: JSON.stringify({ message: message, saveToSentItems: true }),
    });

    // Step 2: Find the message we just sent in engineering-team's Inbox (match by exact subject + sender)
    var foundMessage = null;
    for (var attempt = 0; attempt < FIND_MESSAGE_RETRIES; attempt++) {
      await sleep(attempt === 0 ? FIND_MESSAGE_DELAY_MS : FIND_MESSAGE_RETRY_DELAY_MS);
      var inboxUrl =
        '/users/' + encodeURIComponent(ENGINEERING_TEAM_EMAIL) +
        '/mailFolders/inbox/messages?' +
        '$orderby=receivedDateTime desc&$top=20&$select=id,subject,from,receivedDateTime';
      var result = await AuthService.graphRequest(inboxUrl);
      var messages = (result && result.value) || [];
      for (var i = 0; i < messages.length; i++) {
        var msg = messages[i];
        var fromAddress = msg.from && msg.from.emailAddress && msg.from.emailAddress.address;
        if (msg.subject === outboundSubject && fromAddress === userEmail) {
          foundMessage = msg;
          break;
        }
      }
      if (foundMessage) break;
    }

    if (!foundMessage) {
      showStatus(
        'Could not find the sent message in engineering-team inbox. Ensure you have Full Access to that mailbox.',
        true
      );
      return;
    }

    // Step 3: Send reply from engineering-team to the user (same thread)
    var replyUrl =
      '/users/' + encodeURIComponent(ENGINEERING_TEAM_EMAIL) +
      '/messages/' + encodeURIComponent(foundMessage.id) +
      '/reply';
    await AuthService.graphRequest(replyUrl, {
      method: 'POST',
      body: JSON.stringify({ comment: REPLY_COMMENT }),
    });

    showStatus('Email sent and reply received in the same thread.', false);
  } catch (error) {
    console.error('Send RFQ error:', error);
    var msg = error && error.message ? error.message : String(error);
    var errLower = msg.toLowerCase();
    if (errLower.indexOf('not found') !== -1 || errLower.indexOf('404') !== -1 || errLower.indexOf('default folder') !== -1) {
      msg =
        'Email was sent to engineering-team. To enable the auto-reply, your admin must grant you Full Access (and Send As) to engineering-team@Hexa729.onmicrosoft.com.';
    } else if (msg.indexOf('403') !== -1 || msg.indexOf('Access') !== -1) {
      msg =
        'Access denied. In Azure, add delegated permission Mail.Send.Shared (Microsoft Graph), then sign out and sign in again. Ensure Full Access and Send As (or Send on behalf) for engineering-team@Hexa729.onmicrosoft.com in Exchange.';
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
      body: JSON.stringify({ comment: REPLY_DIRECT_TO_CUSTOMER_COMMENT }),
    });
    replySent = true;

    var replySubject = 'Re: ' + originalSubject;
    var replySubjectLower = replySubject.trim().toLowerCase();
    var originalSubjectLower = originalSubject.trim().toLowerCase();
    var userEmailLower = (userEmail || '').toLowerCase();
    var foundInCustomer1 = null;
    for (var attempt = 0; attempt < CUSTOMER_1_FIND_RETRIES; attempt++) {
      await sleep(attempt === 0 ? CUSTOMER_1_FIND_DELAY_MS : CUSTOMER_1_FIND_RETRY_DELAY_MS);
      var inboxUrl =
        '/users/' + encodeURIComponent(CUSTOMER_1_EMAIL) +
        '/mailFolders/inbox/messages?' +
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
      showStatus(
        'Reply sent. Could not find your reply in customer-1 inbox. Ensure Full Access to customer-1@hexa729.onmicrosoft.com.',
        true
      );
      return;
    }

    var customer1ReplyUrl =
      '/users/' + encodeURIComponent(CUSTOMER_1_EMAIL) +
      '/messages/' + encodeURIComponent(foundInCustomer1.id) +
      '/reply';
    await AuthService.graphRequest(customer1ReplyUrl, {
      method: 'POST',
      body: JSON.stringify({ comment: CUSTOMER_1_AUTO_REPLY_COMMENT }),
    });

    showStatus('Reply sent and automated reply from customer-1 received. Open that reply to send the final reply.', false);
  } catch (error) {
    console.error('Reply direct to customer error:', error);
    var msg = error && error.message ? error.message : String(error);
    var errLower = msg.toLowerCase();
    if (replySent) {
      if (errLower.indexOf('not found') !== -1 || errLower.indexOf('404') !== -1 || errLower.indexOf('default folder') !== -1) {
        showStatus(
          'Reply sent. Automated reply from customer-1 failed: ensure Full Access and Send As (or Send on behalf) for customer-1@hexa729.onmicrosoft.com.',
          true
        );
      } else if (msg.indexOf('403') !== -1 || msg.indexOf('Access') !== -1) {
        showStatus(
          'Reply sent. Automated reply from customer-1 failed: ensure Send As (or Send on behalf) for customer-1@hexa729.onmicrosoft.com, then sign in again.',
          true
        );
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
    showStatus('Original message not found. Cannot reply.', true);
    return;
  }

  var replyBtn = document.getElementById('reply-to-original-btn');
  if (replyBtn) replyBtn.disabled = true;

  var replyToOriginalSent = false;
  try {
    showStatus('Sending reply...', false);

    var originalMsg = await AuthService.graphRequest('/me/messages/' + encodeURIComponent(originalRestId) + '?$select=subject');
    var originalSubject = (originalMsg && originalMsg.subject) ? originalMsg.subject.trim() : '';

    await AuthService.graphRequest('/me/messages/' + encodeURIComponent(originalRestId) + '/reply', {
      method: 'POST',
      body: JSON.stringify({ comment: REPLY_TO_ORIGINAL_COMMENT }),
    });
    replyToOriginalSent = true;

    var replySubject = 'Re: ' + originalSubject;
    var replySubjectLower = replySubject.trim().toLowerCase();
    var originalSubjectLower = originalSubject.trim().toLowerCase();
    var userEmailLower = (userEmail || '').toLowerCase();
    var foundInCustomer1 = null;
    for (var attempt = 0; attempt < CUSTOMER_1_FIND_RETRIES; attempt++) {
      await sleep(attempt === 0 ? CUSTOMER_1_FIND_DELAY_MS : CUSTOMER_1_FIND_RETRY_DELAY_MS);
      var inboxUrl =
        '/users/' + encodeURIComponent(CUSTOMER_1_EMAIL) +
        '/mailFolders/inbox/messages?' +
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
      showStatus(
        'Reply sent to original. Could not find your reply in customer-1 inbox. Ensure Full Access to customer-1@hexa729.onmicrosoft.com.',
        true
      );
      return;
    }

    var customer1ReplyUrl =
      '/users/' + encodeURIComponent(CUSTOMER_1_EMAIL) +
      '/messages/' + encodeURIComponent(foundInCustomer1.id) +
      '/reply';
    await AuthService.graphRequest(customer1ReplyUrl, {
      method: 'POST',
      body: JSON.stringify({ comment: CUSTOMER_1_AUTO_REPLY_COMMENT }),
    });

    showStatus('Reply sent and automated reply from customer-1 received in the same thread.', false);
  } catch (error) {
    console.error('Reply to original error:', error);
    var msg = error && error.message ? error.message : String(error);
    var prefix = replyToOriginalSent ? 'Reply sent to original. Automated reply from customer-1 failed: ' : 'Reply failed: ';
    var errLower = msg.toLowerCase();
    if (replyToOriginalSent) {
      if (errLower.indexOf('not found') !== -1 || errLower.indexOf('404') !== -1 || errLower.indexOf('default folder') !== -1) {
        showStatus(
          'Reply sent to original. Automated reply from customer-1 failed: ensure Full Access and Send As (or Send on behalf) for customer-1@hexa729.onmicrosoft.com.',
          true
        );
      } else if (msg.indexOf('403') !== -1 || msg.indexOf('Access') !== -1) {
        showStatus(
          'Reply sent to original. Automated reply from customer-1 failed: ensure Send As (or Send on behalf) for customer-1@hexa729.onmicrosoft.com, then sign in again.',
          true
        );
      } else {
        showStatus('Reply sent to original. Automated reply from customer-1 failed: ' + msg, true);
      }
    } else {
      showStatus(prefix + msg, true);
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

    showStatus('Sending final reply...', false);
    var btn = document.getElementById('send-final-reply-btn');
    if (btn) btn.disabled = true;

    await AuthService.graphRequest('/me/messages/' + encodeURIComponent(restId) + '/reply', {
      method: 'POST',
      body: JSON.stringify({ comment: FINAL_REPLY_TO_CUSTOMER_COMMENT }),
    });

    showStatus('Final reply sent to customer.', false);
  } catch (error) {
    console.error('Send final reply error:', error);
    showStatus('Final reply failed: ' + (error && error.message ? error.message : String(error)), true);
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
