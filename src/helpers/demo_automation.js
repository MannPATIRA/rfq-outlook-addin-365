/**
 * Demo Email Automation
 * Simulates automated responses from Engineering Team and Client
 * For demo purposes only - uses controlled email addresses
 */

const nodemailer = require('nodemailer');

// Configuration - Replace with your actual email addresses
const CONFIG = {
    // Email account for sending automated responses
    smtp: {
        host: process.env.SMTP_HOST || 'smtp.gmail.com',
        port: process.env.SMTP_PORT || 587,
        secure: false,
        auth: {
            user: process.env.SMTP_USER || 'your-demo-email@gmail.com',
            pass: process.env.SMTP_PASS || 'your-app-password'
        }
    },
    
    // Demo email addresses (should all be controlled by you)
    addresses: {
        engineering: process.env.ENGINEERING_EMAIL || 'engineering@demo.hexa.ai',
        client: process.env.CLIENT_EMAIL || 'client@demo.hexa.ai',
        salesperson: process.env.SALES_EMAIL || 'sales@demo.hexa.ai'
    },
    
    // Response delay in milliseconds (for realistic demo timing)
    responseDelay: 3000
};

// Create transporter
let transporter = null;

async function initializeTransporter() {
    if (!transporter) {
        transporter = nodemailer.createTransport(CONFIG.smtp);
        
        // Verify connection
        try {
            await transporter.verify();
            console.log('✅ Email transporter ready');
        } catch (error) {
            console.error('❌ Email configuration error:', error.message);
        }
    }
    return transporter;
}

/**
 * Send automated engineering response
 */
async function sendEngineeringResponse(originalEmail, questions) {
    const transport = await initializeTransporter();
    
    const engineeringAnswers = generateEngineeringAnswers(questions);
    
    const emailContent = `
Subject: RE: Technical Questions - NRL FBG Arrays RFQ 41260018

Dear Sales Team,

We have reviewed the technical questions from the NRL RFQ and provide the following responses:

${engineeringAnswers.map((a, i) => `
**Question ${i + 1}: ${a.question}**

${a.answer}
`).join('\n---\n')}

Please incorporate these responses in your client communication.

Best regards,
Engineering Team
engionic Femto Gratings GmbH
    `.trim();
    
    const mailOptions = {
        from: CONFIG.addresses.engineering,
        to: CONFIG.addresses.salesperson,
        subject: `RE: Technical Questions - NRL FBG Arrays RFQ 41260018`,
        text: emailContent,
        inReplyTo: originalEmail?.messageId,
        references: originalEmail?.messageId
    };
    
    // Simulate delay
    await new Promise(resolve => setTimeout(resolve, CONFIG.responseDelay));
    
    try {
        const info = await transport.sendMail(mailOptions);
        console.log('✅ Engineering response sent:', info.messageId);
        return { success: true, messageId: info.messageId, content: emailContent };
    } catch (error) {
        console.error('❌ Failed to send engineering response:', error);
        return { success: false, error: error.message };
    }
}

/**
 * Send automated client response with complete specifications
 */
async function sendClientCompleteResponse(originalEmail) {
    const transport = await initializeTransporter();
    
    const emailContent = `
Subject: RE: Clarification & Missing Specifications - NRL FBG Arrays

Dear engionic Femto Gratings Team,

Thank you for the clarifications and engineering expertise. Here are the complete specifications you requested:

**FBG Wavelengths:**
- FBG 1 Wavelength: 1550.39 nm
- FBG 2 Wavelength: 1555.39 nm

**FBG Specifications:**
- Reflectivity: 10% ±4%
- FWHM: 0.09 nm ±0.02 nm
- SLSR: 8 dB minimum
- FBG Length: 12 mm ±2 mm
- FBG Spacing: 50 mm
- First FBG Position: 5000 mm from fiber start

**Additional Requirements:**
- FemtoPlus technology: Yes (as confirmed)
- Apodization: Yes, required for all FBGs
- Spectrum Datasheet: linear format
- Label: on spool

Please proceed with the formal quotation based on these specifications.

We look forward to receiving your offer.

Best regards,
Dr. Sarah Chen
US Naval Research Laboratory
NRL - Materials Science Division
    `.trim();
    
    const mailOptions = {
        from: CONFIG.addresses.client,
        to: CONFIG.addresses.salesperson,
        subject: `RE: Clarification & Missing Specifications - NRL FBG Arrays`,
        text: emailContent,
        inReplyTo: originalEmail?.messageId,
        references: originalEmail?.messageId
    };
    
    // Simulate delay
    await new Promise(resolve => setTimeout(resolve, CONFIG.responseDelay));
    
    try {
        const info = await transport.sendMail(mailOptions);
        console.log('✅ Client complete response sent:', info.messageId);
        return { success: true, messageId: info.messageId, content: emailContent };
    } catch (error) {
        console.error('❌ Failed to send client response:', error);
        return { success: false, error: error.message };
    }
}

/**
 * Generate engineering answers for client questions
 */
function generateEngineeringAnswers(questions) {
    const answerTemplates = {
        apodization: {
            keywords: ['apodization', 'sidelobe', 'side lobe'],
            answer: `Confirmed. We will apply Gaussian apodization to all FBGs in the array, achieving sidelobe suppression levels of <-15dB (typically -18dB to -20dB). The apodization profile will be optimized for the customer's specified reflectivity of 10%. Our apodization process uses a proprietary phase mask design that ensures consistent results across all gratings in the array.

Technical Details:
- Apodization Type: Gaussian profile
- Sidelobe Suppression: <-15dB (typically -18 to -20dB)
- Spectral Shape: Maintained Gaussian-like main peak
- Process: UV inscription with apodized phase mask`
        },
        temperature: {
            keywords: ['temperature', 'polyimide', 'thermal', 'operating'],
            answer: `The polyimide coating (SM1330-E9/125PI fiber from J-Fiber) is rated for continuous operation at temperatures up to 300°C with short-term excursions to 350°C. This has been validated through our internal thermal cycling tests (500 cycles, -40°C to +300°C). 

For naval applications, we recommend our enhanced polyimide option which provides additional moisture resistance suitable for marine environments.

Specifications:
- Continuous Operating Temperature: -40°C to +300°C
- Short-term Peak: 350°C
- Thermal Cycling Validated: 500 cycles
- Humidity Resistance: Class C per IEC 60068-2-78`
        },
        femtoplus: {
            keywords: ['femtoplus', 'femto plus', 'stability', 'technology'],
            answer: `Yes, FemtoPlus technology will be employed for all gratings in this order. This proprietary process provides significant advantages:

1. **Improved Thermal Stability**: Wavelength drift <10pm/°C (significantly better than standard FBGs)
2. **Enhanced Mechanical Durability**: >1000 strain cycles at 1000µε without degradation
3. **Reduced Hydrogen Sensitivity**: Minimal wavelength shift in hydrogen-rich environments
4. **Long-term Stability**: Proven performance over 10+ years in field applications

The FemtoPlus process uses a proprietary high-temperature annealing procedure that eliminates the unstable component of the grating structure.`
        },
        aerospace: {
            keywords: ['aerospace', 'certification', 'documentation', 'compliance', 'mil'],
            answer: `We can provide comprehensive documentation for aerospace compliance:

**Available Certifications & Documentation:**
1. Material certifications per MIL-STD-810H
2. RoHS 3 Directive 2015/863/EU compliance
3. REACH compliance (SVHC-free declaration)
4. Full traceability documentation (lot tracking from raw fiber to finished sensor)
5. Test reports including:
   - Thermal cycling (-55°C to +125°C, 100 cycles)
   - Vibration (random vibration per MIL-STD-810H, Method 514.7)
   - Humidity exposure (85°C/85%RH, 1000 hours)
   - Mechanical shock (40g, 11ms half-sine)

**Quality Management:**
- Manufacturing: ISO 9001:2015 certified
- For AS9100D requirements: Available with 2-week additional lead time

Please specify which certifications are required for your project, and we will include them with the delivery.`
        }
    };
    
    return questions.map(q => {
        let answer = 'We have reviewed this technical requirement and can confirm compliance. Our engineering team will provide detailed specifications with the formal quotation.';
        
        const questionLower = q.question.toLowerCase();
        
        for (const [key, template] of Object.entries(answerTemplates)) {
            if (template.keywords.some(kw => questionLower.includes(kw))) {
                answer = template.answer;
                break;
            }
        }
        
        return {
            question: q.question,
            answer: answer
        };
    });
}

/**
 * Demo workflow runner
 * Runs through the complete demo workflow automatically
 */
async function runDemoWorkflow() {
    console.log('\n🎬 Starting Demo Workflow...\n');
    
    // Step 1: Initial RFQ (manual - user clicks on RFQ email)
    console.log('1️⃣  User opens RFQ email from NRL...');
    console.log('   → Add-in displays extracted specs, missing specs, and questions\n');
    
    await new Promise(resolve => setTimeout(resolve, 2000));
    
    // Step 2: User sends to engineering
    console.log('2️⃣  User clicks "Send to Engineering Team"...');
    console.log('   → Sending engineering inquiry...\n');
    
    const questions = [
        { question: "Can you confirm if apodization will be applied to reduce sidelobe levels below -15dB?" },
        { question: "What is the maximum continuous operating temperature for the polyimide coating?" },
        { question: "Will you be using your FemtoPlus technology for enhanced temperature stability?" },
        { question: "Can you provide documentation for aerospace compliance certification?" }
    ];
    
    await sendEngineeringResponse(null, questions);
    console.log('   ✅ Engineering response sent!\n');
    
    await new Promise(resolve => setTimeout(resolve, 2000));
    
    // Step 3: User sends to client with engineering answers
    console.log('3️⃣  User reviews engineering answers and clicks "Send Reply to Client"...\n');
    
    await new Promise(resolve => setTimeout(resolve, 2000));
    
    // Step 4: Client responds with complete specs
    console.log('4️⃣  Client reply received with complete specifications...');
    await sendClientCompleteResponse(null);
    console.log('   ✅ Client response with complete specs received!\n');
    
    await new Promise(resolve => setTimeout(resolve, 2000));
    
    // Step 5: Generate documents
    console.log('5️⃣  User clicks "Generate Quote Documents"...');
    console.log('   → Generating Excel specification sheet...');
    console.log('   → Generating PDF quote document...');
    console.log('   ✅ Documents generated!\n');
    
    await new Promise(resolve => setTimeout(resolve, 2000));
    
    // Step 6: Send final quote
    console.log('6️⃣  User reviews documents and clicks "Confirm & Send Quote"...');
    console.log('   ✅ Final quote sent to client with attachments!\n');
    
    console.log('═══════════════════════════════════════════════════════');
    console.log('🎉 Demo Workflow Complete!');
    console.log('═══════════════════════════════════════════════════════\n');
}

// Export functions for use in server
module.exports = {
    initializeTransporter,
    sendEngineeringResponse,
    sendClientCompleteResponse,
    generateEngineeringAnswers,
    runDemoWorkflow,
    CONFIG
};

// Run demo if called directly
if (require.main === module) {
    runDemoWorkflow().catch(console.error);
}
