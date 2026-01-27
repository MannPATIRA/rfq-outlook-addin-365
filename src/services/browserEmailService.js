/**
 * Browser-based Email Service using MSAL.js
 * Works with personal Microsoft accounts (@outlook.com, @hotmail.com, etc.)
 */

// MSAL Configuration - will be set from environment
let msalConfig = null;
let msalInstance = null;
let accessToken = null;

// Conversation tracking
const conversations = new Map();

/**
 * Initialize MSAL with client ID
 */
function initializeMSAL(clientId) {
    msalConfig = {
        auth: {
            clientId: clientId,
            authority: 'https://login.microsoftonline.com/consumers', // Personal accounts
            redirectUri: window.location.origin + '/auth-callback.html'
        },
        cache: {
            cacheLocation: 'localStorage',
            storeAuthStateInCookie: true
        }
    };
    
    msalInstance = new msal.PublicClientApplication(msalConfig);
    console.log('✅ MSAL initialized for personal accounts');
    
    // Check if already logged in
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        console.log('📧 Already logged in as:', accounts[0].username);
        return accounts[0];
    }
    return null;
}

/**
 * Login with popup
 */
async function login() {
    const loginRequest = {
        scopes: ['Mail.Send', 'Mail.ReadWrite', 'User.Read']
    };
    
    try {
        const response = await msalInstance.loginPopup(loginRequest);
        console.log('✅ Logged in as:', response.account.username);
        return response.account;
    } catch (error) {
        console.error('Login failed:', error);
        throw error;
    }
}

/**
 * Get access token (will prompt login if needed)
 */
async function getAccessToken() {
    const accounts = msalInstance.getAllAccounts();
    
    if (accounts.length === 0) {
        await login();
    }
    
    const tokenRequest = {
        scopes: ['Mail.Send', 'Mail.ReadWrite'],
        account: msalInstance.getAllAccounts()[0]
    };
    
    try {
        const response = await msalInstance.acquireTokenSilent(tokenRequest);
        accessToken = response.accessToken;
        return accessToken;
    } catch (error) {
        // If silent fails, try popup
        const response = await msalInstance.acquireTokenPopup(tokenRequest);
        accessToken = response.accessToken;
        return accessToken;
    }
}

/**
 * Send email via Microsoft Graph API
 */
async function sendEmail({ to, subject, body, inReplyTo, saveToSentItems = true }) {
    const token = await getAccessToken();
    
    const message = {
        subject: subject,
        body: {
            contentType: 'HTML',
            content: body
        },
        toRecipients: [{ emailAddress: { address: to } }]
    };
    
    // Add reply headers if this is a reply
    if (inReplyTo) {
        message.internetMessageHeaders = [
            { name: 'In-Reply-To', value: inReplyTo },
            { name: 'References', value: inReplyTo }
        ];
    }
    
    const response = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            message: message,
            saveToSentItems: saveToSentItems
        })
    });
    
    if (!response.ok) {
        const error = await response.json();
        throw new Error(error.error?.message || 'Failed to send email');
    }
    
    // Generate a fake message ID (Graph doesn't return one on send)
    const messageId = `<${Date.now()}.${Math.random().toString(36).substr(2)}@outlook.com>`;
    
    return {
        success: true,
        messageId: messageId
    };
}

/**
 * Create a draft email (so we can get its Message-ID)
 */
async function createDraft({ to, subject, body }) {
    const token = await getAccessToken();
    
    const message = {
        subject: subject,
        body: {
            contentType: 'HTML',
            content: body
        },
        toRecipients: [{ emailAddress: { address: to } }]
    };
    
    const response = await fetch('https://graph.microsoft.com/v1.0/me/messages', {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(message)
    });
    
    if (!response.ok) {
        const error = await response.json();
        throw new Error(error.error?.message || 'Failed to create draft');
    }
    
    return await response.json();
}

/**
 * Send a draft email
 */
async function sendDraft(messageId) {
    const token = await getAccessToken();
    
    const response = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${messageId}/send`, {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${token}`
        }
    });
    
    if (!response.ok) {
        const error = await response.json();
        throw new Error(error.error?.message || 'Failed to send draft');
    }
    
    return { success: true };
}

/**
 * Create an email in the inbox (simulating receiving an email)
 * This creates a draft and moves it to inbox to simulate receiving
 */
async function createInboxMessage({ from, subject, body, inReplyTo }) {
    const token = await getAccessToken();
    
    // We'll create a message directly in the inbox folder
    // First, get the inbox folder ID
    const foldersResponse = await fetch('https://graph.microsoft.com/v1.0/me/mailFolders/inbox', {
        headers: { 'Authorization': `Bearer ${token}` }
    });
    const inbox = await foldersResponse.json();
    
    // Create the message
    const message = {
        subject: subject,
        body: {
            contentType: 'HTML',
            content: body
        },
        from: {
            emailAddress: { address: from, name: from.split('@')[0] }
        },
        isRead: false,
        isDraft: false
    };
    
    // Add threading headers
    if (inReplyTo) {
        message.internetMessageHeaders = [
            { name: 'In-Reply-To', value: inReplyTo },
            { name: 'References', value: inReplyTo }
        ];
        // Make subject a reply
        if (!subject.startsWith('Re:')) {
            message.subject = 'Re: ' + subject;
        }
    }
    
    const response = await fetch('https://graph.microsoft.com/v1.0/me/messages', {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(message)
    });
    
    if (!response.ok) {
        const errorData = await response.json();
        console.error('Failed to create message:', errorData);
        throw new Error(errorData.error?.message || 'Failed to create inbox message');
    }
    
    const created = await response.json();
    return {
        success: true,
        messageId: created.internetMessageId || created.id,
        id: created.id
    };
}

/**
 * Full workflow: Send to engineering and simulate reply
 */
async function sendToEngineeringWithReply(rfqData, userEmail) {
    const conversationId = `conv-${Date.now()}`;
    
    // 1. Send email TO engineering
    const subject = `[Engineering Review] RFQ: ${rfqData.customerName} - FBG Specifications`;
    const body = buildEngineeringEmailBody(rfqData);
    
    console.log('📤 Sending to engineering...');
    const sent = await sendEmail({
        to: userEmail, // Send to yourself to see it
        subject: subject,
        body: body
    });
    
    // Store conversation
    conversations.set(conversationId, {
        rfqData: rfqData,
        engineeringMessageId: sent.messageId,
        stage: 'sent_to_engineering'
    });
    
    // 2. After delay, create the "reply" in inbox
    await new Promise(resolve => setTimeout(resolve, 3000));
    
    console.log('📥 Creating engineering reply in inbox...');
    const replyBody = buildEngineeringReplyBody(rfqData);
    
    const reply = await createInboxMessage({
        from: 'engineering@hexa.ai',
        subject: subject,
        body: replyBody,
        inReplyTo: sent.messageId
    });
    
    conversations.get(conversationId).engineeringReplyId = reply.messageId;
    conversations.get(conversationId).stage = 'engineering_replied';
    conversations.get(conversationId).engineeringAnswers = getEngineeringAnswers();
    
    return {
        success: true,
        conversationId: conversationId,
        answers: getEngineeringAnswers()
    };
}

/**
 * Send to client and simulate reply
 */
async function sendToClientWithReply(conversationId, userEmail) {
    const conv = conversations.get(conversationId);
    if (!conv) throw new Error('Conversation not found');
    
    // 1. Send email TO client
    const subject = `Re: RFQ for Fiber Bragg Grating Sensors - Clarification Needed`;
    const body = buildClientEmailBody(conv.rfqData, conv.engineeringAnswers);
    
    console.log('📤 Sending to client...');
    const sent = await sendEmail({
        to: userEmail, // Send to yourself to see it
        subject: subject,
        body: body,
        inReplyTo: conv.engineeringMessageId
    });
    
    conv.clientMessageId = sent.messageId;
    
    // 2. After delay, create the "reply" in inbox
    await new Promise(resolve => setTimeout(resolve, 3000));
    
    console.log('📥 Creating client reply in inbox...');
    const replyBody = buildClientReplyBody(conv.rfqData);
    
    const reply = await createInboxMessage({
        from: 'procurement@nrl.navy.mil',
        subject: subject,
        body: replyBody,
        inReplyTo: sent.messageId
    });
    
    conv.clientReplyId = reply.messageId;
    conv.stage = 'client_replied';
    conv.completedSpecs = getCompletedSpecs();
    
    return {
        success: true,
        completedSpecs: getCompletedSpecs()
    };
}

// ============================================
// Email Body Builders
// ============================================

function buildEngineeringEmailBody(rfqData) {
    return `
        <div style="font-family: Arial, sans-serif; max-width: 800px;">
            <h2>Engineering Review Required</h2>
            <p>Please review the following RFQ and provide technical clarification.</p>
            
            <h3>Customer Information</h3>
            <p><strong>Customer:</strong> ${rfqData.customerName || 'N/A'}</p>
            
            <h3>Specifications Provided</h3>
            <table style="border-collapse: collapse; width: 100%;">
                <tr style="background: #f0f0f0;">
                    <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Parameter</th>
                    <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Value</th>
                </tr>
                ${(rfqData.specsProvided || []).map(spec => `
                    <tr>
                        <td style="border: 1px solid #ddd; padding: 8px;">${spec.name}</td>
                        <td style="border: 1px solid #ddd; padding: 8px;">${spec.value}</td>
                    </tr>
                `).join('')}
            </table>
            
            <h3>Technical Questions from Customer</h3>
            <ol>
                ${(rfqData.questions || []).map(q => `<li>${q.question}</li>`).join('')}
            </ol>
            
            <p style="margin-top: 20px; padding: 10px; background: #fff3cd; border-left: 4px solid #ffc107;">
                Please reply with technical recommendations.
            </p>
        </div>
    `;
}

function buildEngineeringReplyBody(rfqData) {
    const answers = getEngineeringAnswers();
    return `
        <div style="font-family: Arial, sans-serif; max-width: 800px;">
            <p>Hello,</p>
            <p>I've reviewed the RFQ and here are my technical recommendations:</p>
            
            <h3>Answers to Customer Questions</h3>
            ${answers.map((a, i) => `
                <div style="margin: 15px 0; padding: 10px; background: #f8f9fa; border-left: 3px solid #007bff;">
                    <p><strong>Q${i + 1}: ${a.question}</strong></p>
                    <p>${a.answer}</p>
                </div>
            `).join('')}
            
            <h3>Recommended Default Values for Missing Specs</h3>
            <ul>
                <li><strong>Fiber Type:</strong> SMF-28e+ (standard telecom single-mode)</li>
                <li><strong>Coating:</strong> Polyimide (for high-temp applications)</li>
                <li><strong>Connector Type:</strong> FC/APC (angled physical contact)</li>
                <li><strong>FBG Length:</strong> 10mm (standard for strain sensing)</li>
                <li><strong>Reflectivity:</strong> >90%</li>
            </ul>
            
            <p>Best regards,<br>Engineering Team</p>
        </div>
    `;
}

function buildClientEmailBody(rfqData, engineeringAnswers) {
    return `
        <div style="font-family: Arial, sans-serif; max-width: 800px;">
            <p>Dear ${rfqData.customerName || 'Customer'},</p>
            
            <p>Thank you for your RFQ for Fiber Bragg Grating sensors. Our engineering team has reviewed your requirements.</p>
            
            <h3>Technical Clarifications</h3>
            ${(engineeringAnswers || []).map((a, i) => `
                <div style="margin: 10px 0;">
                    <p><strong>${i + 1}. ${a.question}</strong></p>
                    <p style="margin-left: 15px;">${a.answer}</p>
                </div>
            `).join('')}
            
            <h3>Information Required for Quote</h3>
            <p>To provide an accurate quotation, please confirm the following specifications:</p>
            <ol>
                ${(rfqData.specsMissing || []).map(spec => `<li>${spec}</li>`).join('')}
            </ol>
            
            <p>Best regards,<br>Hexa Technologies</p>
        </div>
    `;
}

function buildClientReplyBody(rfqData) {
    return `
        <div style="font-family: Arial, sans-serif;">
            <p>Hello,</p>
            
            <p>Thank you for the detailed clarifications. Here are the specifications you requested:</p>
            
            <ul>
                <li><strong>Fiber Type:</strong> SMF-28e+</li>
                <li><strong>Coating:</strong> Polyimide</li>
                <li><strong>Connector Type:</strong> FC/APC</li>
                <li><strong>FBG Length:</strong> 10mm</li>
                <li><strong>Operating Temperature:</strong> -40°C to 300°C</li>
                <li><strong>Strain Range:</strong> ±5000 µε</li>
                <li><strong>Wavelength Range:</strong> 1545-1555nm</li>
                <li><strong>Quantity:</strong> 10 units</li>
            </ul>
            
            <p>Please proceed with the formal quotation.</p>
            
            <p>Best regards,<br>Dr. Sarah Chen<br>US Naval Research Laboratory</p>
        </div>
    `;
}

function getEngineeringAnswers() {
    return [
        {
            question: 'Can you confirm if apodization will be applied to reduce sidelobe levels below -15dB?',
            answer: 'Confirmed. We will apply Gaussian apodization to all FBGs, achieving sidelobe suppression of <-15dB (typically -18dB to -20dB).'
        },
        {
            question: 'What is the maximum continuous operating temperature for the polyimide coating?',
            answer: 'The polyimide coating is rated for continuous operation up to 300°C with short-term excursions to 350°C.'
        },
        {
            question: 'Will you be using FemtoPlus technology for enhanced temperature stability?',
            answer: 'Yes, FemtoPlus technology will be employed providing improved thermal stability with wavelength drift <10pm/°C.'
        },
        {
            question: 'Can you provide documentation for aerospace compliance certification?',
            answer: 'We can provide material certifications per MIL-STD-810, RoHS/REACH compliance, and full traceability documentation. Our process is ISO 9001:2015 certified.'
        }
    ];
}

function getCompletedSpecs() {
    return {
        fiberType: 'SMF-28e+',
        coating: 'Polyimide',
        connectorType: 'FC/APC',
        fbgLength: '10mm',
        operatingTempMin: '-40°C',
        operatingTempMax: '300°C',
        strainRange: '±5000 µε',
        wavelengthRange: '1545-1555nm',
        quantity: 10
    };
}

// Export for use in taskpane
window.EmailService = {
    initializeMSAL,
    login,
    getAccessToken,
    sendEmail,
    createInboxMessage,
    sendToEngineeringWithReply,
    sendToClientWithReply,
    conversations
};
