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

// RFQ Category definitions for email tagging
const RFQ_CATEGORIES = {
  MISSING_DETAILS: { name: 'RFQ - Missing Details', color: 'preset0' },      // Red
  PENDING_ENGINEERING: { name: 'Pending Engineering', color: 'preset1' },    // Orange
  CLARIFICATION: { name: 'RFQ Clarification', color: 'preset7' },            // Blue
  DETAILS_COMPLETE: { name: 'RFQ - Details Complete', color: 'preset4' }     // Green
};

// CategoryService for managing email categories via Graph API
var CategoryService = {
  categoriesInitialized: false,
  lastCategorizedMessageId: null,
  rfqCategoryNames: null,

  getRfqCategoryNames: function() {
    if (!this.rfqCategoryNames) {
      this.rfqCategoryNames = [];
      for (var key in RFQ_CATEGORIES) {
        this.rfqCategoryNames.push(RFQ_CATEGORIES[key].name);
      }
    }
    return this.rfqCategoryNames;
  },

  async ensureCategoriesExist() {
    if (this.categoriesInitialized) return;
    if (!AuthService.isSignedIn()) return;
    
    try {
      var existingCategories = await AuthService.graphRequest('/me/outlook/masterCategories');
      var existingMap = {};
      (existingCategories.value || []).forEach(function(c) {
        existingMap[c.displayName] = c;
      });
      
      for (var key in RFQ_CATEGORIES) {
        var cat = RFQ_CATEGORIES[key];
        var existing = existingMap[cat.name];
        
        if (!existing) {
          // Category doesn't exist - create it
          try {
            await AuthService.graphRequest('/me/outlook/masterCategories', {
              method: 'POST',
              body: JSON.stringify({
                displayName: cat.name,
                color: cat.color
              })
            });
            console.log('Created category:', cat.name, 'with color:', cat.color);
          } catch (createErr) {
            console.warn('Could not create category ' + cat.name + ':', createErr);
          }
        } else if (existing.color !== cat.color) {
          // Category exists but has wrong color - update it
          try {
            await AuthService.graphRequest('/me/outlook/masterCategories/' + encodeURIComponent(existing.id), {
              method: 'PATCH',
              body: JSON.stringify({
                color: cat.color
              })
            });
            console.log('Updated category color:', cat.name, 'to', cat.color);
          } catch (updateErr) {
            console.warn('Could not update category color ' + cat.name + ':', updateErr);
          }
        }
      }
      this.categoriesInitialized = true;
      console.log('RFQ categories initialized');
    } catch (e) {
      console.warn('Could not initialize categories:', e);
    }
  },

  async setCategory(messageId, categoryName) {
    if (!messageId || !categoryName) return;
    if (!AuthService.isSignedIn()) return;
    
    // Avoid re-categorizing the same message repeatedly
    var cacheKey = messageId + ':' + categoryName;
    if (this.lastCategorizedMessageId === cacheKey) return;
    
    try {
      // First get current categories to preserve non-RFQ categories
      var msgData = await AuthService.graphRequest('/me/messages/' + encodeURIComponent(messageId) + '?$select=categories');
      var currentCategories = (msgData && msgData.categories) || [];
      var rfqNames = this.getRfqCategoryNames();
      
      // Filter out any existing RFQ categories, keep others
      var newCategories = currentCategories.filter(function(c) {
        return rfqNames.indexOf(c) === -1;
      });
      
      // Add the new RFQ category
      newCategories.push(categoryName);
      
      await AuthService.graphRequest('/me/messages/' + encodeURIComponent(messageId), {
        method: 'PATCH',
        body: JSON.stringify({ categories: newCategories })
      });
      this.lastCategorizedMessageId = cacheKey;
      console.log('Set category on message:', categoryName);
    } catch (e) {
      console.warn('Could not set category:', e);
    }
  },

  async removeAllRfqCategories(messageId) {
    if (!messageId) return;
    if (!AuthService.isSignedIn()) return;
    
    try {
      // Get current categories
      var msgData = await AuthService.graphRequest('/me/messages/' + encodeURIComponent(messageId) + '?$select=categories');
      var currentCategories = (msgData && msgData.categories) || [];
      var rfqNames = this.getRfqCategoryNames();
      
      // Filter out RFQ categories, keep others
      var newCategories = currentCategories.filter(function(c) {
        return rfqNames.indexOf(c) === -1;
      });
      
      await AuthService.graphRequest('/me/messages/' + encodeURIComponent(messageId), {
        method: 'PATCH',
        body: JSON.stringify({ categories: newCategories })
      });
    } catch (e) {
      console.warn('Could not remove categories:', e);
    }
  }
};

var currentOriginalMessageRestId = null;
var lastUpdatedConversationId = null;

// Helper function to find and update the original customer email in a conversation thread
async function updateOriginalCustomerEmailCategory(currentMessageId, normalizedSubject) {
  try {
    // Get conversation ID of current message
    var currentMsg = await AuthService.graphRequest('/me/messages/' + encodeURIComponent(currentMessageId) + '?$select=conversationId');
    if (!currentMsg || !currentMsg.conversationId) return;
    
    var conversationId = currentMsg.conversationId;
    
    // Avoid updating the same conversation repeatedly
    if (lastUpdatedConversationId === conversationId) return;
    lastUpdatedConversationId = conversationId;
    
    // Find all messages in this conversation from the customer
    var searchUrl = '/me/messages?$filter=conversationId eq \'' + conversationId + '\'&$select=id,from,subject,receivedDateTime&$orderby=receivedDateTime asc&$top=20';
    var result = await AuthService.graphRequest(searchUrl);
    var messages = (result && result.value) || [];
    
    // Find the original message (first message from customer that's not a reply)
    for (var i = 0; i < messages.length; i++) {
      var msg = messages[i];
      var fromAddr = msg.from && msg.from.emailAddress && msg.from.emailAddress.address;
      if (fromAddr && fromAddr.toLowerCase() === CUSTOMER_1_EMAIL.toLowerCase()) {
        var msgSubject = (msg.subject || '').toLowerCase();
        // Original message won't have "Re:" prefix
        if (msgSubject.indexOf('re:') !== 0) {
          // This is likely the original customer RFQ - update its category
          await CategoryService.setCategory(msg.id, RFQ_CATEGORIES.DETAILS_COMPLETE.name);
          console.log('Updated original customer email category to Details Complete');
          break;
        }
      }
    }
  } catch (e) {
    console.warn('Could not update original customer email category:', e);
  }
}

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
  // Initialize RFQ email categories
  await CategoryService.ensureCategoriesExist();
  renderInitialRfqFromPreset();
  renderEngineeringAnswersFromPreset();
  renderQuoteSummary();
  renderConfirmedDetails();
  setupCollapsibleSections();
  setupAttachmentPreviewButtons();
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
  if (btn && !btn.classList.contains('btn-success')) btn.disabled = !AuthService.isSignedIn();
  var replyBtn = document.getElementById('reply-to-original-btn');
  if (replyBtn && !replyBtn.classList.contains('btn-success')) replyBtn.disabled = !AuthService.isSignedIn();
  var finalReplyBtn = document.getElementById('send-final-reply-btn');
  if (finalReplyBtn && !finalReplyBtn.classList.contains('btn-success')) finalReplyBtn.disabled = !AuthService.isSignedIn();
  var directReplyBtn = document.getElementById('reply-direct-to-customer-btn');
  if (directReplyBtn && !directReplyBtn.classList.contains('btn-success')) directReplyBtn.disabled = !AuthService.isSignedIn();
}

function renderMissingDetails(containerId) {
  var data = typeof RFQ_DATA !== 'undefined' ? RFQ_DATA : null;
  var list = document.getElementById(containerId);
  if (!list || !data || !data.missingDetails || !data.missingDetails.length) return;

  list.innerHTML = data.missingDetails.map(function (m) {
    var badgeClass = m.importance === 'critical' ? 'badge-critical' : 'badge-high';
    return '<li><span class="badge ' + badgeClass + '">' + escapeHtml(m.importance || '') + '</span> ' + escapeHtml(m.field) + '</li>';
  }).join('');
}

function renderQuestionsToAsk(containerId) {
  var data = typeof RFQ_DATA !== 'undefined' ? RFQ_DATA : null;
  var list = document.getElementById(containerId);
  if (!list || !data || !data.questionsToAsk || !data.questionsToAsk.length) return;

  list.innerHTML = data.questionsToAsk.map(function (q) {
    return '<li>' + escapeHtml(q) + '</li>';
  }).join('');
}

function renderInitialRfqFromPreset() {
  var data = typeof RFQ_DATA !== 'undefined' ? RFQ_DATA : null;
  if (!data) return;

  var titleEl = document.getElementById('rfq-summary-title');
  var metaEl = document.getElementById('rfq-summary-meta');
  if (titleEl) titleEl.textContent = 'RFQ #' + (data.technicalSpecs.offerNumber || '41260018') + ' – ' + (data.technicalSpecs.customer || 'NRL');
  if (metaEl) metaEl.textContent = '2 FBG Arrays | Qty: ' + (data.technicalSpecs.quantity || 10);

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
            '<textarea class="accordion-answer-input" data-question-id="' + (q.id || idx) + '" rows="1" placeholder="Edit answer…">' + escapeHtml(q.answer) + '</textarea>' +
          '</div>' +
        '</div>'
      );
    }).join('');
    accordion.querySelectorAll('.accordion-header').forEach(function (btn) {
      btn.addEventListener('click', function () {
        var item = btn.closest('.accordion-item');
        if (item) {
          item.classList.toggle('open');
          var ta = item.querySelector('.accordion-answer-input');
          if (ta) setTimeout(function () { autoResizeAnswerTextarea(ta); }, 0);
        }
      });
    });
    accordion.querySelectorAll('.accordion-answer-input').forEach(function (ta) {
      autoResizeAnswerTextarea(ta);
      ta.addEventListener('input', function () { autoResizeAnswerTextarea(ta); });
    });
  }

  renderMissingDetails('missing-details-list');
  renderQuestionsToAsk('questions-to-ask-list');
}

function stripHtml(html) {
  if (!html) return '';
  var div = document.createElement('div');
  div.innerHTML = html;
  return (div.textContent || div.innerText || '').trim();
}

function renderEngineeringAnswersFromPreset() {
  var templates = typeof EMAIL_TEMPLATES !== 'undefined' ? EMAIL_TEMPLATES : null;
  if (!templates || !templates.engineeringReply || !templates.engineeringReply.body) return;

  var container = document.getElementById('engineering-answers-container');
  if (!container) return;

  var body = templates.engineeringReply.body;
  // Parse proper numbered list items from the HTML content
  var parser = new DOMParser();
  var doc = parser.parseFromString(body, 'text/html');
  var listItems = doc.querySelectorAll('li');
  
  var cards = [];
  if (listItems.length > 0) {
    listItems.forEach(function(li, index) {
      // Clean up the text content
      var text = li.textContent.trim();
      // Try to find a bold title
      var strong = li.querySelector('strong');
      var title = strong ? strong.textContent.replace(/:$/, '') : 'Item ' + (index + 1);
      var content = text;
      
      // If we stripped the title from text, ensure content is correct
      if (strong && text.startsWith(strong.textContent)) {
         content = text.substring(strong.textContent.length).trim();
         // clean up leading colon or space
         content = content.replace(/^[:\s]+/, '');
      }

      cards.push({ title: title, content: content });
    });
  } else {
    // Fallback if no <li> found
    var text = stripHtml(body);
    cards.push({ title: 'Assessment', content: text });
  }

  container.innerHTML = '<ul class="engineering-answers-list">' + cards.map(function (item) {
    return (
      '<li>' +
        '<span class="eng-label">' + escapeHtml(item.title) + '</span>' +
        '<span class="eng-value">' + escapeHtml(item.content) + '</span>' +
      '</li>'
    );
  }).join('') + '</ul>';

  renderMissingDetails('engineering-missing-details-list');
  renderQuestionsToAsk('engineering-questions-to-ask-list');
}

function renderQuoteSummary() {
  var productEl = document.getElementById('quote-product');
  var amountEl = document.getElementById('quote-amount');
  var termsEl = document.getElementById('quote-terms');
  var validEl = document.getElementById('quote-valid');
  var totalEl = document.getElementById('quote-total-value');
  if (productEl) productEl.textContent = '10x 2 FBG arrays, FC/APC both ends';
  if (amountEl) amountEl.textContent = '€2,206.40';
  if (termsEl) termsEl.textContent = 'Payment in advance, EXW';
  if (validEl) validEl.textContent = '30 days';
  if (totalEl) totalEl.textContent = '€2,206.40';
}

/** Preset list of customer-confirmed items (from customer reply with details). */
var CONFIRMED_DETAILS_PRESET = [
  { label: 'Reflectivity', value: '10%' },
  { label: 'Calibration range', value: '-40°C to +85°C' },
  { label: 'Delivery', value: '8 weeks from order confirmation' },
  { label: 'SLSR minimum', value: '8.0 dB (confirmed as acceptance criterion)' },
  { label: 'Calibration data', value: '5-point standard characterisation' },
];

function renderConfirmedDetails() {
  var list = document.getElementById('confirmed-details-list');
  if (!list) return;
  list.innerHTML = CONFIRMED_DETAILS_PRESET.map(function (item) {
    return '<li class="confirmed-detail-item"><span class="confirmed-label">' + escapeHtml(item.label) + '</span><span class="confirmed-value">' + escapeHtml(item.value) + '</span></li>';
  }).join('');
}

function setupCollapsibleSections() {
  function toggleSection(headerId, bodyId) {
    var header = document.getElementById(headerId);
    var body = document.getElementById(bodyId);
    if (!header || !body) return;
    header.addEventListener('click', function () {
      var open = body.style.display !== 'none';
      body.style.display = open ? 'none' : 'block';
      header.setAttribute('aria-expanded', open ? 'false' : 'true');
      var chevron = header.querySelector('.chevron');
      if (chevron) chevron.style.transform = open ? 'rotate(-90deg)' : 'rotate(0deg)';
    });
  }
  toggleSection('toggle-customer-questions', 'body-customer-questions');
  toggleSection('toggle-information-needed', 'body-information-needed');
  toggleSection('toggle-engineering-missing', 'body-engineering-missing');
  toggleSection('toggle-engineering-confirmations', 'body-engineering-confirmations');
}

function setupAttachmentPreviewButtons() {
  var baseUrl = typeof window !== 'undefined' && window.location && window.location.origin ? window.location.origin : '';
  var pdfUrl = baseUrl + '/assets/templates/412600xx.pdf';
  var xlsxUrl = baseUrl + '/assets/templates/412600xx.xlsx';
  document.getElementById('view-pdf-btn')?.addEventListener('click', function () {
    window.open(pdfUrl, '_blank');
  });
  document.getElementById('view-xlsx-btn')?.addEventListener('click', function () {
    window.open(xlsxUrl, '_blank');
  });
}

function escapeHtml(s) {
  if (s == null) return '';
  var div = document.createElement('div');
  div.textContent = s;
  return div.innerHTML;
}

function autoResizeAnswerTextarea(ta) {
  if (!ta || ta.nodeName !== 'TEXTAREA') return;
  ta.style.height = 'auto';
  var h = ta.scrollHeight;
  var minH = 60;
  var maxH = 500;
  ta.style.height = Math.min(Math.max(h, minH), maxH) + 'px';
}

function showContext(contextId) {
  var blocks = ['context-neutral', 'context-initial-rfq', 'context-engineering-reply', 'context-customer-reply-details'];
  blocks.forEach(function (id) {
    var el = document.getElementById(id);
    if (el) el.classList.toggle('visible', id === contextId);
  });
}

async function detectEmailContext() {
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

  // Get REST ID for current message to apply categories
  var currentMessageRestId = null;
  if (item.itemId && Office.context.mailbox.convertToRestId) {
    try {
      currentMessageRestId = Office.context.mailbox.convertToRestId(
        item.itemId,
        Office.MailboxEnums.RestVersion.v2_0
      );
    } catch (e) {
      console.warn('Could not convert item ID to REST ID:', e);
    }
  }

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
    // Apply "RFQ Clarification" category (Blue) to engineering reply
    if (currentMessageRestId && AuthService.isSignedIn()) {
      CategoryService.setCategory(currentMessageRestId, RFQ_CATEGORIES.CLARIFICATION.name);
    }
    return;
  }

  if (isReplyFromCustomer1) {
    showContext('context-customer-reply-details');
    if (statusEl) { statusEl.textContent = ''; statusEl.classList.remove('error', 'success'); }
    // Apply "RFQ - Details Complete" category (Green) to customer reply with details
    if (currentMessageRestId && AuthService.isSignedIn()) {
      CategoryService.setCategory(currentMessageRestId, RFQ_CATEGORIES.DETAILS_COMPLETE.name);
      
      // Also update the original customer email in the thread to remove "Missing Details"
      // Find original message by conversation ID
      updateOriginalCustomerEmailCategory(currentMessageRestId, normalizedSubject);
    }
    return;
  }

  if (isFromCustomer1) {
    showContext('context-initial-rfq');
    if (statusEl) { statusEl.textContent = ''; statusEl.classList.remove('error', 'success'); }
    // Apply "RFQ - Missing Details" category (Red) to initial customer RFQ
    if (currentMessageRestId && AuthService.isSignedIn()) {
      CategoryService.setCategory(currentMessageRestId, RFQ_CATEGORIES.MISSING_DETAILS.name);
    }
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

function setButtonSent(btnId) {
  var btn = document.getElementById(btnId);
  if (!btn) return;
  btn.textContent = 'Sent';
  btn.classList.add('btn-success');
  btn.disabled = true;
}

async function handleSendRfq() {
  if (!AuthService.isSignedIn()) {
    showStatus('Please sign in.', true);
    return;
  }

  var user = AuthService.getUser();
  var userEmail = user ? user.email : '';

  var outboundSubject = makeUniqueSubject();
  var originalMessageRestId = null;
  if (Office.context.mailbox && Office.context.mailbox.item) {
    try {
      var itemId = Office.context.mailbox.item.itemId;
      if (itemId && Office.context.mailbox.convertToRestId) {
        var restId = Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);
        if (restId) {
          saveOriginalForSubject(outboundSubject, restId);
          originalMessageRestId = restId;
        }
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
      body: { contentType: 'HTML', content: getEngineeringReviewBody() },
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
      body: JSON.stringify({ message: { body: { contentType: 'HTML', content: getEngineeringReplyComment() } } }),
    });

    // Update original email category to "Pending Engineering" (Orange)
    if (originalMessageRestId) {
      await CategoryService.setCategory(originalMessageRestId, RFQ_CATEGORIES.PENDING_ENGINEERING.name);
    }

    setButtonSent('send-rfq-btn');
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
      body: JSON.stringify({ message: { body: { contentType: 'HTML', content: getCustomerClarificationComment() } } }),
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
      body: JSON.stringify({ message: { body: { contentType: 'HTML', content: getCustomerReplyWithDetailsComment() } } }),
    });

    setButtonSent('reply-direct-to-customer-btn');
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
      body: JSON.stringify({ message: { body: { contentType: 'HTML', content: getCustomerClarificationComment() } } }),
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
      body: JSON.stringify({ message: { body: { contentType: 'HTML', content: getCustomerReplyWithDetailsComment() } } }),
    });

    setButtonSent('reply-to-original-btn');
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

  var restId;
  try {
    restId = Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);
  } catch (e) {
    showStatus('Could not get message id.', true);
    return;
  }
  if (!restId) {
    showStatus('Could not get message id.', true);
    return;
  }

  var fallback = typeof ATTACHMENTS !== 'undefined' ? ATTACHMENTS : null;
  var baseUrl = typeof window !== 'undefined' && window.location && window.location.origin ? window.location.origin : '';

  var btn = document.getElementById('send-final-reply-btn');
  if (btn) btn.disabled = true;
  showStatus('Creating reply and adding attachments...', false);

  try {
    var pdfAtt = fallback && fallback.pdf ? fallback.pdf : null;
    var xlsxAtt = fallback && fallback.xlsx ? fallback.xlsx : null;
    if (baseUrl) {
      try {
        var pdfRes = await fetch(baseUrl + '/api/attachments/pdf');
        if (pdfRes && pdfRes.ok) pdfAtt = await pdfRes.json();
      } catch (e) { /* use fallback */ }
      try {
        var xlsxRes = await fetch(baseUrl + '/api/attachments/xlsx');
        if (xlsxRes && xlsxRes.ok) xlsxAtt = await xlsxRes.json();
      } catch (e) { /* use fallback */ }
    }
    if (!pdfAtt || !xlsxAtt || !pdfAtt.contentBytes || !xlsxAtt.contentBytes) {
      showStatus('Attachment data not loaded. Ensure 412600xx.pdf and 412600xx.xlsx are in src/assets/templates.', true);
      if (btn) btn.disabled = !AuthService.isSignedIn();
      return;
    }

    var createReplyResult = await AuthService.graphRequest('/me/messages/' + encodeURIComponent(restId) + '/createReply', {
      method: 'POST',
    });
    var draftId = createReplyResult && createReplyResult.id;
    if (!draftId) {
      showStatus('Could not create reply draft.', true);
      if (btn) btn.disabled = !AuthService.isSignedIn();
      return;
    }

    await AuthService.graphRequest('/me/messages/' + encodeURIComponent(draftId) + '/attachments', {
      method: 'POST',
      body: JSON.stringify({
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: pdfAtt.name,
        contentType: pdfAtt.contentType,
        contentBytes: pdfAtt.contentBytes,
      }),
    });
    await AuthService.graphRequest('/me/messages/' + encodeURIComponent(draftId) + '/attachments', {
      method: 'POST',
      body: JSON.stringify({
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: xlsxAtt.name,
        contentType: xlsxAtt.contentType,
        contentBytes: xlsxAtt.contentBytes,
      }),
    });

    var htmlBody = getFinalQuoteComment();
    await AuthService.graphRequest('/me/messages/' + encodeURIComponent(draftId), {
      method: 'PATCH',
      body: JSON.stringify({ body: { contentType: 'HTML', content: htmlBody } }),
    });

    await AuthService.graphRequest('/me/messages/' + encodeURIComponent(draftId) + '/send', {
      method: 'POST',
    });

    setButtonSent('send-final-reply-btn');
    showStatus('Final quote sent to customer with PDF and XLSX attachments.', false);
  } catch (error) {
    console.error('Send final reply error:', error);
    showStatus('Final quote failed: ' + (error && error.message ? error.message : String(error)), true);
    var finalBtn = document.getElementById('send-final-reply-btn');
    if (finalBtn) finalBtn.disabled = !AuthService.isSignedIn();
  }
}

function sleep(ms) {
  return new Promise(function (resolve) {
    setTimeout(resolve, ms);
  });
}
