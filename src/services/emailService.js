/**
 * Email Service - Microsoft Graph API Integration
 * Handles real email sending and automated reply simulation
 */

require('dotenv').config();
require('isomorphic-fetch');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');

// In-memory store for conversation state (use Redis/DB in production)
const conversationStore = new Map();

// MSAL Configuration
const msalConfig = {
  auth: {
    clientId: process.env.AZURE_CLIENT_ID,
    clientSecret: process.env.AZURE_CLIENT_SECRET,
    authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`
  }
};

let msalClient = null;
let graphClient = null;

/**
 * Initialize the Graph API client
 */
async function initializeGraphClient() {
  if (!process.env.AZURE_CLIENT_ID || !process.env.AZURE_CLIENT_SECRET || !process.env.AZURE_TENANT_ID) {
    console.warn('⚠️  Azure credentials not configured. Email service will run in demo mode.');
    return null;
  }

  try {
    msalClient = new ConfidentialClientApplication(msalConfig);
    
    // Create Graph client with token acquisition
    graphClient = Client.init({
      authProvider: async (done) => {
        try {
          const result = await msalClient.acquireTokenByClientCredential({
            scopes: ['https://graph.microsoft.com/.default']
          });
          done(null, result.accessToken);
        } catch (error) {
          done(error, null);
        }
      }
    });

    console.log('✅ Microsoft Graph client initialized');
    return graphClient;
  } catch (error) {
    console.error('❌ Failed to initialize Graph client:', error.message);
    return null;
  }
}

/**
 * Generate a unique Message-ID
 */
function generateMessageId() {
  const timestamp = Date.now();
  const random = Math.random().toString(36).substring(2, 15);
  const domain = process.env.EMAIL_DOMAIN || 'hexa.ai';
  return `<${timestamp}.${random}@${domain}>`;
}

/**
 * Send an email via Microsoft Graph API
 */
async function sendEmail({ to, cc, subject, body, inReplyTo, references, conversationId, fromDisplayName }) {
  const messageId = generateMessageId();
  
  // Build the email message
  const message = {
    subject: subject,
    body: {
      contentType: 'HTML',
      content: body
    },
    toRecipients: (Array.isArray(to) ? to : [to]).map(email => ({
      emailAddress: { address: email }
    })),
    // Custom headers for threading
    internetMessageHeaders: []
  };

  // Add CC if provided
  if (cc) {
    message.ccRecipients = (Array.isArray(cc) ? cc : [cc]).map(email => ({
      emailAddress: { address: email }
    }));
  }

  // Add threading headers
  if (inReplyTo) {
    message.internetMessageHeaders.push({
      name: 'In-Reply-To',
      value: inReplyTo
    });
  }
  
  if (references) {
    message.internetMessageHeaders.push({
      name: 'References',
      value: references
    });
  }

  // If Graph client is available, send real email
  if (graphClient) {
    try {
      const userEmail = process.env.USER_EMAIL;
      const result = await graphClient
        .api(`/users/${userEmail}/sendMail`)
        .post({ message, saveToSentItems: true });
      
      console.log(`✅ Email sent: ${subject}`);
      
      return {
        success: true,
        messageId: messageId,
        conversationId: conversationId || messageId
      };
    } catch (error) {
      console.error('❌ Failed to send email:', error.message);
      throw error;
    }
  } else {
    // Demo mode - just log and return mock response
    console.log(`📧 [DEMO] Would send email:`);
    console.log(`   To: ${to}`);
    console.log(`   Subject: ${subject}`);
    console.log(`   In-Reply-To: ${inReplyTo || 'none'}`);
    
    return {
      success: true,
      messageId: messageId,
      conversationId: conversationId || messageId,
      demo: true
    };
  }
}

/**
 * Send email to engineering team
 */
async function sendToEngineering({ rfqData, userEmail, originalMessageId }) {
  const conversationId = `conv-${Date.now()}`;
  const subject = `[Engineering Review] RFQ: ${rfqData.customerName || 'Customer'} - FBG Specifications`;
  
  const body = `
    <div style="font-family: Arial, sans-serif; max-width: 800px;">
      <h2>Engineering Review Required</h2>
      <p>Please review the following RFQ and provide technical clarification.</p>
      
      <h3>Customer Information</h3>
      <p><strong>Customer:</strong> ${rfqData.customerName || 'N/A'}</p>
      <p><strong>Organization:</strong> ${rfqData.organization || 'N/A'}</p>
      
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
      
      <h3>Missing Specifications</h3>
      <ul>
        ${(rfqData.specsMissing || []).map(spec => `<li>${spec}</li>`).join('')}
      </ul>
      
      <h3>Technical Questions from Customer</h3>
      <ol>
        ${(rfqData.questions || []).map(q => `<li>${q.question}</li>`).join('')}
      </ol>
      
      <p style="margin-top: 20px; padding: 10px; background: #fff3cd; border-left: 4px solid #ffc107;">
        Please reply to this email with your technical recommendations and answers to the customer's questions.
      </p>
    </div>
  `;

  const result = await sendEmail({
    to: process.env.ENGINEERING_EMAIL || 'engineering@hexa.ai',
    subject: subject,
    body: body,
    conversationId: conversationId
  });

  // Store conversation state for threading
  conversationStore.set(conversationId, {
    originalMessageId: originalMessageId,
    engineeringMessageId: result.messageId,
    rfqData: rfqData,
    userEmail: userEmail,
    stage: 'awaiting_engineering',
    createdAt: new Date()
  });

  return {
    ...result,
    conversationId: conversationId
  };
}

/**
 * Simulate engineering reply (sends real email to user's inbox)
 */
async function simulateEngineeringReply({ conversationId, delayMs = 3000 }) {
  const conversation = conversationStore.get(conversationId);
  if (!conversation) {
    throw new Error('Conversation not found');
  }

  // Wait for specified delay to simulate response time
  await new Promise(resolve => setTimeout(resolve, delayMs));

  const rfqData = conversation.rfqData;
  const subject = `Re: [Engineering Review] RFQ: ${rfqData.customerName || 'Customer'} - FBG Specifications`;
  
  // Generate technical answers
  const engineeringAnswers = [
    {
      question: rfqData.questions?.[0]?.question || 'Fiber coating compatibility',
      answer: 'For the specified temperature range up to 300°C, we recommend our high-temperature polyimide coating. Standard acrylate coating is limited to 85°C and would not be suitable.'
    },
    {
      question: rfqData.questions?.[1]?.question || 'Wavelength tolerance',
      answer: 'We can achieve ±0.1nm center wavelength tolerance for the 1550nm range. Tighter tolerances of ±0.05nm are available at additional cost.'
    },
    {
      question: rfqData.questions?.[2]?.question || 'Reflectivity specifications',
      answer: 'Standard reflectivity is >90% which is suitable for most sensing applications. We can provide >99% reflectivity for applications requiring stronger signal return.'
    },
    {
      question: rfqData.questions?.[3]?.question || 'Lead time',
      answer: 'Standard lead time is 4-6 weeks for custom FBG arrays. Expedited delivery in 2-3 weeks is possible with a 25% surcharge.'
    }
  ];

  const body = `
    <div style="font-family: Arial, sans-serif; max-width: 800px;">
      <p>Hello,</p>
      <p>I've reviewed the RFQ and here are my technical recommendations:</p>
      
      <h3>Answers to Customer Questions</h3>
      ${engineeringAnswers.map((a, i) => `
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
      
      <p>Please clarify the operating temperature range with the customer, as this will determine the final coating selection.</p>
      
      <p>Best regards,<br>Engineering Team</p>
    </div>
  `;

  const result = await sendEmail({
    to: conversation.userEmail || process.env.USER_EMAIL,
    subject: subject,
    body: body,
    inReplyTo: conversation.engineeringMessageId,
    references: conversation.engineeringMessageId,
    conversationId: conversationId
  });

  // Update conversation state
  conversation.stage = 'engineering_replied';
  conversation.engineeringReplyId = result.messageId;
  conversation.engineeringAnswers = engineeringAnswers;
  conversationStore.set(conversationId, conversation);

  return {
    ...result,
    answers: engineeringAnswers
  };
}

/**
 * Send email to client with questions
 */
async function sendToClient({ conversationId, engineeringAnswers }) {
  const conversation = conversationStore.get(conversationId);
  if (!conversation) {
    throw new Error('Conversation not found');
  }

  const rfqData = conversation.rfqData;
  const subject = `Re: RFQ for Fiber Bragg Grating Sensors - Clarification Needed`;

  const body = `
    <div style="font-family: Arial, sans-serif; max-width: 800px;">
      <p>Dear ${rfqData.customerName || 'Customer'},</p>
      
      <p>Thank you for your RFQ for Fiber Bragg Grating sensors. Our engineering team has reviewed your requirements and we have some clarifications below, along with a few questions to ensure we provide an accurate quote.</p>
      
      <h3>Technical Clarifications</h3>
      ${(engineeringAnswers || conversation.engineeringAnswers || []).map((a, i) => `
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
      
      <p>Please reply to this email with the requested information and we will prepare your formal quotation.</p>
      
      <p>Best regards,<br>
      Hexa Technologies<br>
      quotes@hexa.ai</p>
    </div>
  `;

  const result = await sendEmail({
    to: rfqData.customerEmail || process.env.DEMO_CLIENT_EMAIL,
    subject: subject,
    body: body,
    inReplyTo: conversation.originalMessageId,
    references: conversation.originalMessageId,
    conversationId: conversationId
  });

  // Update conversation state
  conversation.stage = 'awaiting_client';
  conversation.clientMessageId = result.messageId;
  conversationStore.set(conversationId, conversation);

  return result;
}

/**
 * Simulate client reply with complete specifications
 */
async function simulateClientReply({ conversationId, delayMs = 3000 }) {
  const conversation = conversationStore.get(conversationId);
  if (!conversation) {
    throw new Error('Conversation not found');
  }

  await new Promise(resolve => setTimeout(resolve, delayMs));

  const rfqData = conversation.rfqData;
  const subject = `Re: Re: RFQ for Fiber Bragg Grating Sensors - Clarification Needed`;

  const completedSpecs = {
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

  const body = `
    <div style="font-family: Arial, sans-serif;">
      <p>Hello,</p>
      
      <p>Thank you for the detailed clarifications. Here are the specifications you requested:</p>
      
      <ul>
        <li><strong>Fiber Type:</strong> ${completedSpecs.fiberType}</li>
        <li><strong>Coating:</strong> ${completedSpecs.coating}</li>
        <li><strong>Connector Type:</strong> ${completedSpecs.connectorType}</li>
        <li><strong>FBG Length:</strong> ${completedSpecs.fbgLength}</li>
        <li><strong>Operating Temperature:</strong> ${completedSpecs.operatingTempMin} to ${completedSpecs.operatingTempMax}</li>
        <li><strong>Strain Range:</strong> ${completedSpecs.strainRange}</li>
        <li><strong>Wavelength Range:</strong> ${completedSpecs.wavelengthRange}</li>
        <li><strong>Quantity:</strong> ${completedSpecs.quantity} units</li>
      </ul>
      
      <p>Please proceed with the formal quotation at your earliest convenience.</p>
      
      <p>Best regards,<br>
      ${rfqData.customerName || 'Customer'}<br>
      ${rfqData.organization || ''}</p>
    </div>
  `;

  const result = await sendEmail({
    to: conversation.userEmail || process.env.USER_EMAIL,
    subject: subject,
    body: body,
    inReplyTo: conversation.clientMessageId,
    references: `${conversation.originalMessageId} ${conversation.clientMessageId}`,
    conversationId: conversationId
  });

  // Update conversation state
  conversation.stage = 'client_replied';
  conversation.clientReplyId = result.messageId;
  conversation.completedSpecs = completedSpecs;
  conversationStore.set(conversationId, conversation);

  return {
    ...result,
    completedSpecs: completedSpecs
  };
}

/**
 * Get conversation state
 */
function getConversation(conversationId) {
  return conversationStore.get(conversationId);
}

/**
 * List all active conversations
 */
function listConversations() {
  return Array.from(conversationStore.entries()).map(([id, conv]) => ({
    id,
    stage: conv.stage,
    customer: conv.rfqData?.customerName,
    createdAt: conv.createdAt
  }));
}

module.exports = {
  initializeGraphClient,
  sendEmail,
  sendToEngineering,
  simulateEngineeringReply,
  sendToClient,
  simulateClientReply,
  getConversation,
  listConversations
};
