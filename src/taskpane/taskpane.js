/**
 * Hexa RFQ Manager - Outlook Add-in
 * Technical Sales RFQ Management System for Fiber Bragg Grating Products
 */

// ============================================
// Configuration & State Management
// ============================================

const CONFIG = {
    // Azure App Client ID - SET THIS!
    azureClientId: 'YOUR_AZURE_CLIENT_ID_HERE', // <-- Replace with your App ID
    
    // Email addresses for demo (configure these)
    engineeringEmail: 'engineering@demo.hexa.ai',
    clientEmail: 'client@demo.hexa.ai',
    
    // Company info
    companyName: 'engionic Femto Gratings GmbH',
    companyAddress: 'Am Stollen 19b, 38640 Goslar, Germany',
    
    // Offer details
    offerNumber: '41260018',
    unitPrice: 220.64,
    customsDeclaration: 60.00
};

// Application state
let appState = {
    currentView: 'initial',
    emailContext: null,
    originalRfqMessageId: null,
    extractedSpecs: {},
    missingSpecs: [],
    clientQuestions: [],
    engineeringAnswers: [],
    completeSpecs: {},
    workflowStep: 1,
    userEmail: null,
    isLoggedIn: false
};

// Store the conversation ID for threading
let currentConversationId = null;

// ============================================
// Demo Data - Simulated Email Content
// ============================================

const DEMO_DATA = {
    // Initial RFQ Email Content (simulated)
    rfqEmail: {
        subject: "RFQ: FBG Arrays for Naval Research Project",
        from: "procurement@nrl.navy.mil",
        body: `Dear engionic Femto Gratings,

We are requesting a quotation for Fiber Bragg Grating (FBG) arrays for our Naval Research Laboratory project.

SPECIFICATIONS:
- Quantity: 10 pieces
- Configuration: 2 FBG arrays per fiber
- Connector Type: FC/APC on both ends
- Fiber Type: SM1330-E9/125PI
- Total Fiber Length: 10,050 mm
- Number of FBGs: 2

TECHNICAL QUESTIONS:
1. Can you confirm if apodization will be applied to reduce sidelobe levels below -15dB?
2. What is the maximum continuous operating temperature for the polyimide coating?
3. Will you be using your FemtoPlus technology for enhanced temperature stability?
4. Can you provide documentation for aerospace compliance certification?

Please provide a formal quotation at your earliest convenience.

Best regards,
Dr. Sarah Chen
US Naval Research Laboratory
NRL - Materials Science Division`
    },
    
    // Extracted specifications from RFQ
    extractedSpecs: {
        'Customer': 'NRL - US Naval Research Laboratory',
        'Quantity': '10 pcs',
        'Configuration': '2 FBG arrays with FC/APC connectors on both ends',
        'Fiber Type': 'SM1330-E9/125PI',
        'Connector Type': 'FC/APC both ends',
        'Total Fiber Length': '10,050 mm',
        'Number of FBGs': '2',
        'Fiber Coating': 'Polyimide'
    },
    
    // Missing specifications
    missingSpecs: [
        { name: 'FBG Wavelength', description: 'Specific wavelength for each FBG (nm)', required: true },
        { name: 'Reflectivity', description: 'Target reflectivity percentage (%)', required: true },
        { name: 'FWHM', description: 'Full Width Half Maximum (nm)', required: true },
        { name: 'SLSR', description: 'Minimum Side Lobe Suppression Ratio (dB)', required: true },
        { name: 'FBG Length', description: 'Physical length of each FBG (mm)', required: true },
        { name: 'FBG Spacing', description: 'Spacing between gratings (mm)', required: true },
        { name: 'FBG Position', description: 'Position of first FBG from fiber start (mm)', required: false }
    ],
    
    // Client questions extracted
    clientQuestions: [
        {
            id: 1,
            question: "Can you confirm if apodization will be applied to reduce sidelobe levels below -15dB?",
            aiAnswer: "Yes, we can apply Gaussian apodization to all FBGs in the array. This technique shapes the refractive index profile to achieve sidelobe suppression levels typically better than -15dB. The apodization profile is optimized for your specified reflectivity to maintain spectral characteristics while minimizing sidelobe levels."
        },
        {
            id: 2,
            question: "What is the maximum continuous operating temperature for the polyimide coating?",
            aiAnswer: "The SM1330-E9/125PI fiber with polyimide coating is rated for continuous operation at temperatures up to 300°C, with short-term excursions possible up to 350°C. This has been validated through extensive thermal cycling tests. For your specific application requirements, we recommend discussing the operating environment with our engineering team."
        },
        {
            id: 3,
            question: "Will you be using your FemtoPlus technology for enhanced temperature stability?",
            aiAnswer: "Yes, we offer FemtoPlus technology which provides improved thermal stability with wavelength drift typically <10pm/°C and enhanced mechanical durability. This is our recommended approach for applications requiring high temperature stability and long-term reliability."
        },
        {
            id: 4,
            question: "Can you provide documentation for aerospace compliance certification?",
            aiAnswer: "We can provide material certifications and test reports suitable for aerospace applications. Our manufacturing process is ISO 9001 certified, and we maintain full traceability documentation. Specific aerospace certifications depend on the applicable standards for your project. Please confirm which certifications are required."
        }
    ],
    
    // Engineering team response (simulated)
    engineeringResponse: {
        subject: "RE: Technical Questions - NRL FBG Arrays RFQ 41260018",
        answers: [
            {
                questionId: 1,
                answer: "Confirmed. We will apply Gaussian apodization to all FBGs in the array, achieving sidelobe suppression levels of <-15dB (typically -18dB to -20dB). The apodization profile will be optimized for the customer's specified reflectivity of 10%. Our apodization process uses a proprietary phase mask design that ensures consistent results across all gratings in the array."
            },
            {
                questionId: 2,
                answer: "The polyimide coating (SM1330-E9/125PI fiber from J-Fiber) is rated for continuous operation at temperatures up to 300°C with short-term excursions to 350°C. This has been validated through our internal thermal cycling tests (500 cycles, -40°C to +300°C). For naval applications, we recommend our enhanced polyimide option which provides additional moisture resistance."
            },
            {
                questionId: 3,
                answer: "Yes, FemtoPlus technology will be employed for all gratings in this order. This provides: (1) Improved thermal stability with wavelength drift <10pm/°C, (2) Enhanced mechanical durability with >1000 strain cycles without degradation, (3) Reduced hydrogen sensitivity. The FemtoPlus process uses a proprietary high-temperature annealing procedure that ensures long-term stability."
            },
            {
                questionId: 4,
                answer: "We can provide the following documentation for aerospace compliance: (1) Material certifications per MIL-STD-810, (2) RoHS and REACH compliance certificates, (3) Full traceability documentation from raw materials to finished product, (4) Test reports for thermal cycling, vibration, and humidity exposure. Our manufacturing process is ISO 9001:2015 certified. For specific AS9100 requirements, please allow 2 additional weeks lead time."
            }
        ]
    },
    
    // Client response with complete specs (simulated)
    clientCompleteResponse: {
        subject: "RE: Clarification & Missing Specifications - NRL FBG Arrays",
        body: `Thank you for the clarifications and engineering expertise. Here are the complete specifications:

FBG 1 Wavelength: 1550.39 nm
FBG 2 Wavelength: 1555.39 nm (calculated with 5nm spacing)
Reflectivity: 10% ±4%
FWHM: 0.09 nm ±0.02 nm
SLSR: 8 dB minimum
FBG Length: 12 mm ±2 mm
FBG Spacing: 50 mm
First FBG Position: 5000 mm from fiber start

Additional notes:
- Please use FemtoPlus technology as confirmed
- Apodization is required for all FBGs
- Label: on spool
- Spectrum datasheet format: linear

Please proceed with the quotation.

Best regards,
Dr. Sarah Chen`,
        completeSpecs: {
            'Customer': 'NRL - US Naval Research Laboratory',
            'Customer Number': '20797',
            'Quantity': '10 pcs',
            'Configuration': '2 FBG arrays with FC/APC connectors on both ends',
            'Fiber Type': 'SM1330-E9/125PI',
            'Connector Type': 'FC/APC both ends',
            'Total Fiber Length': '10,050 mm',
            'Number of FBGs': '2',
            'Fiber Coating': 'Polyimide',
            'FBG 1 Wavelength': '1550.39 nm',
            'FBG 2 Wavelength': '1555.39 nm',
            'Wavelength Tolerance': '±0.1 nm',
            'Reflectivity': '10%',
            'Reflectivity Tolerance': '±4%',
            'FWHM': '0.09 nm',
            'FWHM Tolerance': '±0.02 nm',
            'SLSR Minimum': '8 dB',
            'FBG Length': '12 mm',
            'FBG Length Tolerance': '±2 mm',
            'FBG Spacing': '50 mm',
            'First FBG Position': '5000 mm',
            'FemtoPlus': 'Yes',
            'Apodized': 'Yes',
            'Spectrum Datasheet': 'linear',
            'Label': 'on spool'
        }
    }
};

// ============================================
// Office.js Initialization
// ============================================

Office.onReady((info) => {
    // Initialize MSAL for email sending
    initializeEmailService();
    
    if (info.host === Office.HostType.Outlook) {
        initializeAddin();
    } else {
        // For testing outside Outlook
        console.log('Running in standalone mode');
        initializeStandaloneDemo();
    }
});

async function initializeEmailService() {
    if (typeof window.EmailService !== 'undefined') {
        // Initialize MSAL with your Azure Client ID
        const account = window.EmailService.initializeMSAL(CONFIG.azureClientId);
        if (account) {
            appState.userEmail = account.username;
            appState.isLoggedIn = true;
            console.log('✅ Already logged in as:', account.username);
        }
    } else {
        console.warn('Email service not loaded');
    }
}

async function ensureLoggedIn() {
    if (!appState.isLoggedIn && typeof window.EmailService !== 'undefined') {
        try {
            const account = await window.EmailService.login();
            appState.userEmail = account.username;
            appState.isLoggedIn = true;
            showToast('✅ Logged in as ' + account.username);
        } catch (error) {
            showError('Please login to send emails: ' + error.message);
            throw error;
        }
    }
}

function initializeAddin() {
    // Get the current email item
    const item = Office.context.mailbox.item;
    
    if (item) {
        analyzeEmail(item);
    } else {
        showError('No email selected');
    }
}

function initializeStandaloneDemo() {
    // For demo/testing without Outlook
    setTimeout(() => {
        processInitialRfq();
    }, 1000);
}

// ============================================
// Email Analysis
// ============================================

async function analyzeEmail(item) {
    showLoading(true, 'Analyzing email...');
    
    try {
        // Get email details
        const subject = item.subject;
        const from = item.from?.emailAddress || '';
        
        // Get email body
        item.body.getAsync(Office.CoercionType.Text, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const body = result.value;
                
                // Store original message ID for threading
                appState.originalRfqMessageId = item.itemId;
                
                // Determine email context
                const context = determineEmailContext(subject, from, body);
                appState.emailContext = context;
                
                // Update UI based on context
                handleEmailContext(context, { subject, from, body });
            } else {
                showError('Failed to read email body');
            }
            
            showLoading(false);
        });
    } catch (error) {
        showLoading(false);
        showError('Error analyzing email: ' + error.message);
    }
}

function determineEmailContext(subject, from, body) {
    // Check if this is an engineering team reply
    if (from.includes(CONFIG.engineeringEmail) || 
        subject.includes('RE: Technical Questions') ||
        body.includes('engineering response') ||
        body.includes('Engineering Team')) {
        return 'engineering_reply';
    }
    
    // Check if this is a client reply with complete specs
    if (body.includes('complete specifications') ||
        body.includes('Please proceed with the quotation') ||
        (body.includes('Wavelength') && body.includes('Reflectivity') && body.includes('FWHM') && body.includes('SLSR'))) {
        return 'client_complete';
    }
    
    // Check if this is a new RFQ
    if (subject.toLowerCase().includes('rfq') || 
        subject.toLowerCase().includes('request for quote') ||
        body.toLowerCase().includes('requesting a quotation')) {
        return 'initial_rfq';
    }
    
    // Default to initial RFQ for demo
    return 'initial_rfq';
}

function handleEmailContext(context, emailData) {
    updateContextDisplay(context);
    
    switch (context) {
        case 'initial_rfq':
            processInitialRfq();
            break;
        case 'engineering_reply':
            processEngineeringReply();
            break;
        case 'client_complete':
            processClientComplete();
            break;
        default:
            processInitialRfq();
    }
}

function updateContextDisplay(context) {
    const display = document.getElementById('emailContextDisplay');
    const contextLabels = {
        'initial_rfq': '📧 New RFQ Detected',
        'engineering_reply': '⚙️ Engineering Response',
        'client_complete': '✅ Complete Specs Received'
    };
    display.textContent = contextLabels[context] || 'Processing...';
}

// ============================================
// View Processing Functions
// ============================================

function processInitialRfq() {
    appState.currentView = 'initial';
    appState.workflowStep = 1;
    appState.extractedSpecs = DEMO_DATA.extractedSpecs;
    appState.missingSpecs = DEMO_DATA.missingSpecs;
    appState.clientQuestions = DEMO_DATA.clientQuestions;
    
    updateWorkflowIndicator(1);
    renderExtractedSpecs();
    renderMissingSpecs();
    renderClientQuestions();
    
    switchView('initialView');
    showLoading(false);
}

function processEngineeringReply() {
    appState.currentView = 'engineering_reply';
    appState.workflowStep = 2;
    appState.engineeringAnswers = DEMO_DATA.engineeringResponse.answers;
    
    updateWorkflowIndicator(2);
    renderEngineeringAnswers();
    renderStillMissingSpecs();
    
    switchView('engineeringReplyView');
    showLoading(false);
}

function processClientComplete() {
    appState.currentView = 'client_complete';
    appState.workflowStep = 4;
    appState.completeSpecs = DEMO_DATA.clientCompleteResponse.completeSpecs;
    
    updateWorkflowIndicator(4);
    renderCompleteSpecs();
    
    switchView('clientReplyView');
    showLoading(false);
}

// ============================================
// UI Rendering Functions
// ============================================

function renderExtractedSpecs() {
    const container = document.getElementById('specsProvidedBody');
    const count = document.getElementById('specsProvidedCount');
    const specs = appState.extractedSpecs;
    
    let html = '';
    let specCount = 0;
    
    for (const [label, value] of Object.entries(specs)) {
        html += `
            <div class="spec-item">
                <span class="spec-label">${label}</span>
                <span class="spec-value">${value}</span>
            </div>
        `;
        specCount++;
    }
    
    container.innerHTML = html;
    count.textContent = specCount;
}

function renderMissingSpecs() {
    const container = document.getElementById('specsMissingBody');
    const count = document.getElementById('specsMissingCount');
    const specs = appState.missingSpecs;
    
    let html = '';
    
    specs.forEach(spec => {
        html += `
            <div class="spec-item">
                <div>
                    <span class="spec-label">${spec.name}</span>
                    ${spec.required ? '<span class="badge bg-danger ms-1" style="font-size:9px;">Required</span>' : ''}
                    <div style="font-size:11px;color:#888;">${spec.description}</div>
                </div>
            </div>
        `;
    });
    
    container.innerHTML = html;
    count.textContent = specs.length;
}

function renderClientQuestions() {
    const container = document.getElementById('questionsBody');
    const count = document.getElementById('questionsCount');
    const questions = appState.clientQuestions;
    
    let html = '';
    
    questions.forEach((q, index) => {
        html += `
            <div class="question-item">
                <div class="question-text">"${q.question}"</div>
                <div class="ai-answer">
                    <div class="ai-answer-label">
                        <i class="bi bi-robot me-1"></i>AI Suggested Answer
                    </div>
                    ${q.aiAnswer}
                </div>
                <div class="mt-2 d-flex gap-2">
                    <button class="btn btn-sm btn-outline-primary" onclick="sendSingleToEngineering(${q.id})">
                        <i class="bi bi-gear me-1"></i>Ask Engineering
                    </button>
                    <button class="btn btn-sm btn-outline-success" onclick="approveAiAnswer(${q.id})">
                        <i class="bi bi-check me-1"></i>Use AI Answer
                    </button>
                </div>
            </div>
        `;
    });
    
    container.innerHTML = html;
    count.textContent = questions.length;
}

function renderEngineeringAnswers() {
    const container = document.getElementById('engineeringAnswersBody');
    const answers = appState.engineeringAnswers;
    const questions = appState.clientQuestions;
    
    let html = '';
    
    answers.forEach((a) => {
        const question = questions.find(q => q.id === a.questionId);
        if (question) {
            html += `
                <div class="question-item">
                    <div class="question-text">"${question.question}"</div>
                    <div class="engineering-answer">
                        <div class="engineering-answer-label">
                            <i class="bi bi-gear-fill me-1"></i>Engineering Team Response
                        </div>
                        ${a.answer}
                    </div>
                </div>
            `;
        }
    });
    
    container.innerHTML = html;
}

function renderStillMissingSpecs() {
    const container = document.getElementById('stillMissingBody');
    const count = document.getElementById('stillMissingCount');
    // Same as missing specs at this stage
    const specs = appState.missingSpecs;
    
    let html = '';
    
    specs.forEach(spec => {
        html += `
            <div class="spec-item">
                <div>
                    <span class="spec-label">${spec.name}</span>
                    <div style="font-size:11px;color:#888;">${spec.description}</div>
                </div>
            </div>
        `;
    });
    
    container.innerHTML = html;
    count.textContent = specs.length;
}

function renderCompleteSpecs() {
    const container = document.getElementById('completeSpecsBody');
    const specs = appState.completeSpecs;
    
    let html = '<div style="max-height:300px;overflow-y:auto;">';
    
    // Group specs by category
    const generalSpecs = ['Customer', 'Customer Number', 'Quantity', 'Configuration'];
    const fiberSpecs = ['Fiber Type', 'Connector Type', 'Total Fiber Length', 'Fiber Coating'];
    const fbgSpecs = ['Number of FBGs', 'FBG 1 Wavelength', 'FBG 2 Wavelength', 'Wavelength Tolerance', 
                      'Reflectivity', 'Reflectivity Tolerance', 'FWHM', 'FWHM Tolerance', 
                      'SLSR Minimum', 'FBG Length', 'FBG Length Tolerance', 'FBG Spacing', 'First FBG Position'];
    const additionalSpecs = ['FemtoPlus', 'Apodized', 'Spectrum Datasheet', 'Label'];
    
    const renderGroup = (title, keys) => {
        let groupHtml = `<div class="mb-3"><strong style="font-size:11px;color:#666;">${title}</strong>`;
        keys.forEach(key => {
            if (specs[key]) {
                groupHtml += `
                    <div class="spec-item">
                        <span class="spec-label">${key}</span>
                        <span class="spec-value">${specs[key]}</span>
                    </div>
                `;
            }
        });
        groupHtml += '</div>';
        return groupHtml;
    };
    
    html += renderGroup('General', generalSpecs);
    html += renderGroup('Fiber Specifications', fiberSpecs);
    html += renderGroup('FBG Specifications', fbgSpecs);
    html += renderGroup('Additional', additionalSpecs);
    
    html += '</div>';
    container.innerHTML = html;
}

// ============================================
// Action Functions
// ============================================

async function sendToEngineering() {
    showLoading(true, 'Preparing to send...');
    
    try {
        // Ensure user is logged in
        await ensureLoggedIn();
        
        showLoading(true, 'Sending to Engineering Team...');
        
        // Build RFQ data from current state
        const rfqData = {
            customerName: appState.extractedSpecs['Customer'] || 'Customer',
            organization: appState.extractedSpecs['Customer'],
            customerEmail: appState.extractedSpecs['Customer Email'] || CONFIG.clientEmail,
            specsProvided: Object.entries(appState.extractedSpecs).map(([name, value]) => ({ name, value })),
            specsMissing: appState.missingSpecs.map(s => s.name),
            questions: appState.clientQuestions
        };
        
        // Use browser-based email service
        if (typeof window.EmailService !== 'undefined') {
            const result = await window.EmailService.sendToEngineeringWithReply(
                rfqData, 
                appState.userEmail
            );
            
            if (result.success) {
                currentConversationId = result.conversationId;
                
                // Store the engineering answers
                appState.engineeringAnswers = result.answers.map((a, i) => ({
                    questionId: appState.clientQuestions[i]?.id || i + 1,
                    answer: a.answer
                }));
                
                showToast('⚙️ Engineering reply received in your inbox!');
                processEngineeringReply();
            }
        } else {
            // Fallback to server API
            const response = await fetch('/api/email/send-to-engineering', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    rfqData: rfqData,
                    userEmail: appState.userEmail,
                    originalMessageId: appState.originalRfqMessageId
                })
            });
            
            const result = await response.json();
            
            if (result.success) {
                currentConversationId = result.conversationId;
                showToast('📨 Email sent to Engineering Team');
                
                // Now trigger the simulated engineering reply
                showLoading(true, 'Waiting for Engineering reply...');
                
                const replyResponse = await fetch('/api/email/simulate-engineering-reply', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        conversationId: currentConversationId,
                        delayMs: 3000
                    })
                });
                
                const replyResult = await replyResponse.json();
                
                if (replyResult.success) {
                    appState.engineeringAnswers = replyResult.answers.map((a, i) => ({
                        questionId: appState.clientQuestions[i]?.id || i + 1,
                        answer: a.answer
                    }));
                    
                    showToast('⚙️ Engineering reply received in your inbox!');
                    processEngineeringReply();
                } else {
                    showError('Failed to get engineering reply: ' + replyResult.error);
                }
            } else {
                showError('Failed to send email: ' + result.error);
            }
        }
    } catch (error) {
        console.error('Error in sendToEngineering:', error);
        showError('Error: ' + error.message);
    }
    
    showLoading(false);
}

function composeEngineeringEmail() {
    const questions = appState.clientQuestions;
    const specs = appState.extractedSpecs;
    
    let body = `Subject: Technical Questions - ${specs['Customer']} FBG Arrays RFQ ${CONFIG.offerNumber}\n\n`;
    body += `Dear Engineering Team,\n\n`;
    body += `We have received an RFQ from ${specs['Customer']} for FBG arrays. `;
    body += `The client has asked several technical questions that require your expertise.\n\n`;
    body += `RFQ DETAILS:\n`;
    for (const [key, value] of Object.entries(specs)) {
        body += `- ${key}: ${value}\n`;
    }
    body += `\nCLIENT QUESTIONS:\n`;
    questions.forEach((q, i) => {
        body += `\n${i + 1}. ${q.question}\n`;
    });
    body += `\nPlease provide detailed technical answers for each question.\n\n`;
    body += `Best regards,\nSales Team`;
    
    return body;
}

async function sendClarificationToClient() {
    showLoading(true, 'Sending to Client...');
    
    try {
        // If we don't have a conversation yet, create one
        if (!currentConversationId) {
            // First send to engineering to create the conversation
            const rfqData = {
                customerName: appState.extractedSpecs['Customer'] || 'Customer',
                organization: appState.extractedSpecs['Customer'],
                customerEmail: appState.extractedSpecs['Customer Email'] || CONFIG.clientEmail,
                specsProvided: Object.entries(appState.extractedSpecs).map(([name, value]) => ({ name, value })),
                specsMissing: appState.missingSpecs.map(s => s.name),
                questions: appState.clientQuestions
            };
            
            const engResponse = await fetch('/api/email/send-to-engineering', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ rfqData, originalMessageId: appState.originalRfqMessageId })
            });
            
            const engResult = await engResponse.json();
            if (engResult.success) {
                currentConversationId = engResult.conversationId;
            }
        }
        
        // Now send to client
        const response = await fetch('/api/email/send-to-client', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                conversationId: currentConversationId,
                engineeringAnswers: appState.clientQuestions.map(q => ({
                    question: q.question,
                    answer: q.aiAnswer
                }))
            })
        });
        
        const result = await response.json();
        
        if (result.success) {
            showToast('📨 Clarification email sent to Client');
            
            // Trigger simulated client reply
            showLoading(true, 'Waiting for Client reply...');
            
            const replyResponse = await fetch('/api/email/simulate-client-reply', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    conversationId: currentConversationId,
                    delayMs: 3000
                })
            });
            
            const replyResult = await replyResponse.json();
            
            if (replyResult.success) {
                appState.completeSpecs = {
                    ...appState.extractedSpecs,
                    ...replyResult.completedSpecs
                };
                showToast('📬 Client reply received with complete specs!');
                processClientComplete();
            }
        } else {
            showError('Failed to send email: ' + result.error);
        }
    } catch (error) {
        console.error('Error in sendClarificationToClient:', error);
        showError('Error: ' + error.message);
    }
    
    showLoading(false);
}

async function sendReplyToClientWithEngineering() {
    showLoading(true, 'Preparing to send...');
    
    try {
        // Ensure user is logged in
        await ensureLoggedIn();
        
        if (!currentConversationId) {
            showError('No conversation ID - please send to engineering first');
            showLoading(false);
            return;
        }
        
        showLoading(true, 'Sending reply to Client...');
        
        // Use browser-based email service
        if (typeof window.EmailService !== 'undefined') {
            const result = await window.EmailService.sendToClientWithReply(
                currentConversationId,
                appState.userEmail
            );
            
            if (result.success) {
                appState.workflowStep = 3;
                updateWorkflowIndicator(3);
                
                // Store the complete specs from client
                appState.completeSpecs = {
                    ...appState.extractedSpecs,
                    ...result.completedSpecs,
                    'Fiber Type': appState.extractedSpecs['Fiber Type'] || result.completedSpecs.fiberType,
                    'Coating': result.completedSpecs.coating,
                    'Connector Type': appState.extractedSpecs['Connector Type'] || result.completedSpecs.connectorType,
                    'FBG Length': result.completedSpecs.fbgLength,
                    'Operating Temp Min': result.completedSpecs.operatingTempMin,
                    'Operating Temp Max': result.completedSpecs.operatingTempMax,
                    'Strain Range': result.completedSpecs.strainRange,
                    'Wavelength Range': result.completedSpecs.wavelengthRange,
                    'Quantity': result.completedSpecs.quantity + ' pcs'
                };
                
                showToast('📬 Client reply received with complete specs!');
                processClientComplete();
            }
        } else {
            // Fallback to server API
            const response = await fetch('/api/email/send-to-client', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    conversationId: currentConversationId,
                    engineeringAnswers: appState.engineeringAnswers
                })
            });
            
            const result = await response.json();
            
            if (result.success) {
                showToast('📨 Reply sent to Client (threaded to original RFQ)');
                appState.workflowStep = 3;
                updateWorkflowIndicator(3);
                
                showLoading(true, 'Waiting for Client reply...');
                
                const replyResponse = await fetch('/api/email/simulate-client-reply', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        conversationId: currentConversationId,
                        delayMs: 3000
                    })
                });
                
                const replyResult = await replyResponse.json();
                
                if (replyResult.success) {
                    appState.completeSpecs = {
                        ...appState.extractedSpecs,
                        ...replyResult.completedSpecs
                    };
                    
                    showToast('📬 Client reply received with complete specs!');
                    processClientComplete();
                } else {
                    showError('Failed to get client reply: ' + replyResult.error);
                }
            } else {
                showError('Failed to send email: ' + result.error);
            }
        }
    } catch (error) {
        console.error('Error in sendReplyToClientWithEngineering:', error);
        showError('Error: ' + error.message);
    }
    
    showLoading(false);
}

function composeClarificationEmail(includeEngineering) {
    const specs = appState.extractedSpecs;
    const missing = appState.missingSpecs;
    const questions = appState.clientQuestions;
    const engAnswers = appState.engineeringAnswers;
    
    let body = `Dear ${specs['Customer']},\n\n`;
    body += `Thank you for your RFQ for FBG arrays. `;
    
    if (includeEngineering) {
        body += `Our engineering team has reviewed your technical questions and provided the following responses:\n\n`;
        body += `TECHNICAL CLARIFICATIONS:\n`;
        engAnswers.forEach((a, i) => {
            const q = questions.find(q => q.id === a.questionId);
            if (q) {
                body += `\nQ${i + 1}: ${q.question}\n`;
                body += `A: ${a.answer}\n`;
            }
        });
    } else {
        body += `Please find our responses to your technical questions below:\n\n`;
        questions.forEach((q, i) => {
            body += `\nQ${i + 1}: ${q.question}\n`;
            body += `A: ${q.aiAnswer}\n`;
        });
    }
    
    body += `\n\nADDITIONAL INFORMATION REQUIRED:\n`;
    body += `To prepare an accurate quotation, we require the following specifications:\n\n`;
    missing.forEach((spec, i) => {
        body += `${i + 1}. ${spec.name}: ${spec.description}\n`;
    });
    
    body += `\nPlease provide these details at your earliest convenience.\n\n`;
    body += `Best regards,\n${CONFIG.companyName}`;
    
    return body;
}

async function generateQuoteDocuments() {
    showLoading(true, 'Generating quote documents...');
    
    try {
        const response = await fetch('/api/generate-documents', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                specifications: appState.completeSpecs,
                offerNumber: CONFIG.offerNumber
            })
        });
        
        const result = await response.json();
        
        if (result.success) {
            // Store document paths for download
            appState.documents = result.documents;
            appState.sessionId = result.sessionId;
            
            appState.workflowStep = 5;
            updateWorkflowIndicator(5);
            switchView('quoteReadyView');
            
            // Update the UI with download links
            const docSection = document.querySelector('#quoteReadyView .section-body');
            if (docSection) {
                docSection.innerHTML = `
                    <div class="mb-3">
                        <p class="text-success"><i class="bi bi-check-circle me-2"></i>Documents generated successfully!</p>
                        <div class="d-grid gap-2">
                            <a href="${result.documents.excel}" class="btn btn-outline-success" download>
                                <i class="bi bi-file-earmark-excel me-2"></i>Download Excel Spec Sheet
                            </a>
                            <a href="${result.documents.pdf}" class="btn btn-outline-danger" download>
                                <i class="bi bi-file-earmark-pdf me-2"></i>Download PDF Quote
                            </a>
                        </div>
                    </div>
                `;
            }
            
            showToast('📄 Quote documents generated!');
        } else {
            showError('Failed to generate documents: ' + result.error);
        }
    } catch (error) {
        console.error('Error generating documents:', error);
        showError('Error: ' + error.message);
    }
    
    showLoading(false);
}

async function sendFinalQuote() {
    showLoading(true, 'Sending final quote to Client...');
    
    const emailBody = composeFinalQuoteEmail();
    
    // In real implementation, this would attach the generated documents
    // and send the email threaded to the original RFQ
    
    setTimeout(() => {
        showToast('✅ Quote sent successfully!');
        showLoading(false);
        
        // Show completion message
        document.querySelector('#quoteReadyView .section-body').innerHTML = `
            <div class="text-center py-4">
                <i class="bi bi-check-circle-fill text-success" style="font-size:48px;"></i>
                <h5 class="mt-3">Quote Sent Successfully!</h5>
                <p class="text-muted">
                    Offer No. ${CONFIG.offerNumber} has been sent to<br>
                    ${appState.completeSpecs['Customer']}
                </p>
                <div class="alert alert-info mt-3" style="font-size:12px;">
                    <strong>Total Value:</strong> €2,266.40<br>
                    <strong>Payment Terms:</strong> Payment in advance<br>
                    <strong>Incoterms:</strong> EXW
                </div>
            </div>
        `;
        
        document.querySelector('#quoteReadyView .action-buttons').innerHTML = `
            <button class="btn btn-hexa btn-hexa-secondary" onclick="resetWorkflow()">
                <i class="bi bi-arrow-repeat me-2"></i>Process New RFQ
            </button>
        `;
    }, 2000);
}

function composeFinalQuoteEmail() {
    const specs = appState.completeSpecs;
    
    let body = `Dear ${specs['Customer']},\n\n`;
    body += `Thank you for providing the additional specifications. `;
    body += `Please find attached our formal quotation (Offer No. ${CONFIG.offerNumber}) `;
    body += `and detailed specification sheet.\n\n`;
    body += `QUOTATION SUMMARY:\n`;
    body += `- Quantity: ${specs['Quantity']}\n`;
    body += `- Description: ${specs['Configuration']}\n`;
    body += `- Unit Price: €${CONFIG.unitPrice.toFixed(2)}\n`;
    body += `- Total Product Value: €${(10 * CONFIG.unitPrice).toFixed(2)}\n`;
    body += `- Customs Declaration: €${CONFIG.customsDeclaration.toFixed(2)}\n`;
    body += `- Total Value: €${(10 * CONFIG.unitPrice + CONFIG.customsDeclaration).toFixed(2)}\n\n`;
    body += `TERMS:\n`;
    body += `- Payment: Payment in advance\n`;
    body += `- Incoterms 2020: EXW\n`;
    body += `- Validity: 30 days\n\n`;
    body += `Please return the signed specification sheet with your order.\n\n`;
    body += `We look forward to your confirmation.\n\n`;
    body += `Best regards,\n${CONFIG.companyName}\n${CONFIG.companyAddress}`;
    
    return body;
}

function downloadDocuments() {
    showToast('📥 Downloading documents...');
    // In real implementation, trigger download of generated files
}

function previewDocument(type) {
    if (type === 'excel') {
        showToast('📊 Opening Excel preview...');
    } else {
        showToast('📄 Opening PDF preview...');
    }
}

function resetWorkflow() {
    appState = {
        currentView: 'initial',
        emailContext: null,
        originalRfqMessageId: null,
        extractedSpecs: {},
        missingSpecs: [],
        clientQuestions: [],
        engineeringAnswers: [],
        completeSpecs: {},
        workflowStep: 1
    };
    processInitialRfq();
}

function sendSingleToEngineering(questionId) {
    showToast(`📨 Question ${questionId} sent to Engineering`);
}

function approveAiAnswer(questionId) {
    showToast(`✅ AI answer approved for question ${questionId}`);
}

// ============================================
// UI Helper Functions
// ============================================

function switchView(viewId) {
    const views = ['initialView', 'engineeringReplyView', 'clientReplyView', 'quoteReadyView'];
    views.forEach(v => {
        document.getElementById(v).classList.remove('active');
    });
    document.getElementById(viewId).classList.add('active');
}

function updateWorkflowIndicator(step) {
    for (let i = 1; i <= 5; i++) {
        const stepEl = document.getElementById(`step${i}`);
        stepEl.classList.remove('active', 'completed');
        
        if (i < step) {
            stepEl.classList.add('completed');
        } else if (i === step) {
            stepEl.classList.add('active');
        }
    }
}

function toggleSection(header) {
    const body = header.nextElementSibling;
    const isCollapsed = header.classList.contains('collapsed');
    
    if (isCollapsed) {
        header.classList.remove('collapsed');
        body.style.maxHeight = body.scrollHeight + 'px';
    } else {
        header.classList.add('collapsed');
        body.style.maxHeight = '0';
    }
}

function showLoading(show, message = 'Loading...') {
    const overlay = document.getElementById('loadingOverlay');
    if (show) {
        overlay.querySelector('p').textContent = message;
        overlay.style.display = 'flex';
    } else {
        overlay.style.display = 'none';
    }
}

function showToast(message) {
    const container = document.getElementById('toastContainer');
    const toast = document.createElement('div');
    toast.className = 'hexa-toast';
    toast.innerHTML = message;
    container.appendChild(toast);
    
    setTimeout(() => {
        toast.style.opacity = '0';
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}

function showError(message) {
    showToast(`❌ ${message}`);
}

// Initialize on page load (for standalone testing)
document.addEventListener('DOMContentLoaded', () => {
    // Check if Office.js is loaded
    if (typeof Office === 'undefined') {
        console.log('Office.js not loaded, running in demo mode');
        setTimeout(() => {
            processInitialRfq();
        }, 1000);
    }
});
