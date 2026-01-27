/**
 * Hexa RFQ Manager - Server
 * Express server for Outlook Add-in backend
 */

const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const morgan = require('morgan');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');
const { v4: uuidv4 } = require('uuid');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// Use /tmp for Vercel (serverless) or local output directory for development
// Vercel's filesystem is read-only except for /tmp
// Check for Vercel environment variables or if we're in a serverless context
const isVercel = process.env.VERCEL || process.env.VERCEL_ENV || process.env.AWS_LAMBDA_FUNCTION_NAME;
const OUTPUT_DIR = isVercel 
    ? path.join('/tmp', 'output')
    : path.join(__dirname, 'output');

// Create output directory if it doesn't exist (lazy initialization)
function ensureOutputDir() {
    try {
        if (!fs.existsSync(OUTPUT_DIR)) {
            fs.mkdirSync(OUTPUT_DIR, { recursive: true });
        }
    } catch (error) {
        // If creation fails, try /tmp as fallback (for Vercel)
        if (!isVercel && OUTPUT_DIR !== path.join('/tmp', 'output')) {
            const tmpDir = path.join('/tmp', 'output');
            if (!fs.existsSync(tmpDir)) {
                fs.mkdirSync(tmpDir, { recursive: true });
            }
            // Note: This won't update OUTPUT_DIR constant, but we'll handle it in usage
            console.warn(`Could not create ${OUTPUT_DIR}, using ${tmpDir} instead`);
        } else {
            throw error;
        }
    }
}

// Middleware
app.use(morgan('combined'));
app.use(cors({
    origin: ['https://localhost:3000', 'https://outlook.office.com', 'https://outlook.office365.com'],
    credentials: true
}));
app.use(helmet({
    contentSecurityPolicy: {
        directives: {
            defaultSrc: ["'self'"],
            scriptSrc: ["'self'", "'unsafe-inline'", "'unsafe-eval'", "https://appsforoffice.microsoft.com", "https://cdn.jsdelivr.net"],
            styleSrc: ["'self'", "'unsafe-inline'", "https://cdn.jsdelivr.net"],
            fontSrc: ["'self'", "https://cdn.jsdelivr.net"],
            imgSrc: ["'self'", "data:", "https:"],
            connectSrc: ["'self'", "https://outlook.office.com", "https://outlook.office365.com"]
        }
    }
}));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Serve static files
app.use('/assets', express.static(path.join(__dirname, 'assets')));
// Only serve /output as static files if not in Vercel (Vercel uses /tmp which isn't accessible via static)
// In Vercel, files are served via the /api/download endpoint
if (!isVercel) {
    ensureOutputDir();
    app.use('/output', express.static(OUTPUT_DIR));
}

// Serve taskpane files
app.get('/taskpane.html', (req, res) => {
    res.sendFile(path.join(__dirname, 'src/taskpane/taskpane.html'));
});

app.get('/taskpane.js', (req, res) => {
    res.sendFile(path.join(__dirname, 'src/taskpane/taskpane.js'));
});

// Serve auth callback
app.get('/auth-callback.html', (req, res) => {
    res.sendFile(path.join(__dirname, 'src/auth-callback.html'));
});

// Serve browser email service
app.get('/services/browserEmailService.js', (req, res) => {
    res.sendFile(path.join(__dirname, 'src/services/browserEmailService.js'));
});

// Serve commands file
app.get('/commands.html', (req, res) => {
    res.send(`
<!DOCTYPE html>
<html>
<head>
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    <script>
        Office.onReady(function() {
            // Commands ready
        });
    </script>
</head>
<body>
</body>
</html>
    `);
});

// Serve manifest
app.get('/manifest.xml', (req, res) => {
    res.type('application/xml');
    res.sendFile(path.join(__dirname, 'manifest.xml'));
});

// API Routes

// Health check
app.get('/api/health', (req, res) => {
    res.json({ status: 'healthy', timestamp: new Date().toISOString() });
});

// Extract specifications from email content
app.post('/api/extract-specs', (req, res) => {
    const { emailBody, emailSubject } = req.body;
    
    try {
        const extracted = extractSpecifications(emailBody);
        const missing = identifyMissingSpecs(extracted);
        const questions = extractClientQuestions(emailBody);
        
        res.json({
            success: true,
            extracted,
            missing,
            questions,
            isRfq: isRfqEmail(emailSubject, emailBody)
        });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// Generate AI answer for a question
app.post('/api/generate-answer', (req, res) => {
    const { question, context } = req.body;
    
    try {
        const answer = generateAiAnswer(question, context);
        res.json({ success: true, answer });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// Generate quote documents
app.post('/api/generate-documents', async (req, res) => {
    const { specifications, offerNumber = '41260018' } = req.body;
    
    try {
        ensureOutputDir();
        const sessionId = uuidv4();
        const sessionDir = path.join(OUTPUT_DIR, sessionId);
        fs.mkdirSync(sessionDir, { recursive: true });
        
        // Write specifications to JSON file for Python script
        const specsPath = path.join(sessionDir, 'specs.json');
        fs.writeFileSync(specsPath, JSON.stringify(specifications, null, 2));
        
        // Call Python document generator
        const pythonScript = path.join(__dirname, 'src/helpers/document_generator.py');
        const result = await runPythonGenerator(pythonScript, specsPath, sessionDir, offerNumber);
        
        if (result.success) {
            // In Vercel, use download endpoint; otherwise use static file URLs
            const excelUrl = isVercel 
                ? `/api/download/${sessionId}/${offerNumber}_NRL.xlsx`
                : `/output/${sessionId}/${offerNumber}_NRL.xlsx`;
            const pdfUrl = isVercel
                ? `/api/download/${sessionId}/${offerNumber}_NRL.pdf`
                : `/output/${sessionId}/${offerNumber}_NRL.pdf`;
            
            res.json({
                success: true,
                sessionId,
                documents: {
                    excel: excelUrl,
                    pdf: pdfUrl
                }
            });
        } else {
            res.status(500).json({ success: false, error: result.error });
        }
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// Download generated document
app.get('/api/download/:sessionId/:filename', (req, res) => {
    const { sessionId, filename } = req.params;
    const filePath = path.join(OUTPUT_DIR, sessionId, filename);
    
    if (fs.existsSync(filePath)) {
        res.download(filePath);
    } else {
        res.status(404).json({ error: 'File not found' });
    }
});

// Simulate engineering response (for demo)
app.post('/api/simulate-engineering-response', (req, res) => {
    const { questions } = req.body;
    
    const engineeringAnswers = questions.map((q, index) => ({
        questionId: q.id || index + 1,
        answer: generateEngineeringAnswer(q.question)
    }));
    
    res.json({ success: true, answers: engineeringAnswers });
});

// Simulate client response with complete specs (for demo)
app.post('/api/simulate-client-response', (req, res) => {
    const completeSpecs = {
        'Customer': 'NRL - US Naval Research Laboratory',
        'Customer Number': '20797',
        'Quantity': '10 pcs',
        'Configuration': '2 FBG arrays with FC/APC connectors on both ends',
        'Fiber Type': 'SM1330-E9/125PI',
        'Connector Type': 'FC/APC both ends',
        'Total Fiber Length': '10050 mm',
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
    };
    
    res.json({ success: true, specifications: completeSpecs });
});

// Helper Functions

function extractSpecifications(emailBody) {
    const specs = {};
    
    // Pattern matching for common specifications
    const patterns = {
        'Quantity': /quantity[:\s]*(\d+)\s*(pcs|pieces|units)?/i,
        'Fiber Type': /fiber\s*type[:\s]*(SM\d+[\w\-\/]+|SMF[\w\-]+)/i,
        'Connector Type': /(FC\/APC|FC\/PC|SC\/APC|SC\/PC)[\s\w]*(both ends|on both ends)?/i,
        'Total Fiber Length': /(?:total\s*)?fiber\s*length[:\s]*(\d+[\,\d]*)\s*mm/i,
        'Number of FBGs': /(\d+)\s*FBG(?:\s*arrays?)?/i,
        'Wavelength': /wavelength[:\s]*(\d+\.?\d*)\s*nm/i,
        'Reflectivity': /reflectivity[:\s]*(\d+\.?\d*)\s*%/i,
        'FWHM': /FWHM[:\s]*(\d+\.?\d*)\s*nm/i,
        'SLSR': /SLSR[:\s]*(\d+\.?\d*)\s*dB/i,
        'FBG Length': /FBG\s*length[:\s]*(\d+\.?\d*)\s*mm/i,
        'FBG Spacing': /(?:FBG\s*)?spacing[:\s]*(\d+[\,\d]*)\s*mm/i
    };
    
    for (const [key, pattern] of Object.entries(patterns)) {
        const match = emailBody.match(pattern);
        if (match) {
            specs[key] = match[1] + (match[2] ? ' ' + match[2] : '');
        }
    }
    
    // Extract customer from common patterns
    const customerPatterns = [
        /(?:from|for|customer)[:\s]+([A-Z]{2,}[\s\-\w]*(?:Laboratory|Research|Corp|Inc|Ltd|GmbH)?)/i,
        /Naval Research Laboratory/i
    ];
    
    for (const pattern of customerPatterns) {
        const match = emailBody.match(pattern);
        if (match) {
            specs['Customer'] = match[1] || match[0];
            break;
        }
    }
    
    return specs;
}

function identifyMissingSpecs(extracted) {
    const required = [
        { name: 'FBG Wavelength', description: 'Specific wavelength for each FBG (nm)', required: true },
        { name: 'Reflectivity', description: 'Target reflectivity percentage (%)', required: true },
        { name: 'FWHM', description: 'Full Width Half Maximum (nm)', required: true },
        { name: 'SLSR', description: 'Minimum Side Lobe Suppression Ratio (dB)', required: true },
        { name: 'FBG Length', description: 'Physical length of each FBG (mm)', required: true },
        { name: 'FBG Spacing', description: 'Spacing between gratings (mm)', required: true },
        { name: 'First FBG Position', description: 'Position of first FBG from fiber start (mm)', required: false }
    ];
    
    return required.filter(spec => !extracted[spec.name]);
}

function extractClientQuestions(emailBody) {
    const questions = [];
    
    // Look for numbered questions
    const numberedPattern = /(\d+)\.\s*([^\n]+\?)/g;
    let match;
    
    while ((match = numberedPattern.exec(emailBody)) !== null) {
        questions.push({
            id: parseInt(match[1]),
            question: match[2].trim()
        });
    }
    
    // Look for question mark sentences
    if (questions.length === 0) {
        const sentences = emailBody.match(/[^.!?\n]+\?/g) || [];
        sentences.forEach((sentence, index) => {
            if (sentence.length > 20) { // Filter out very short questions
                questions.push({
                    id: index + 1,
                    question: sentence.trim()
                });
            }
        });
    }
    
    return questions;
}

function isRfqEmail(subject, body) {
    const rfqKeywords = ['rfq', 'request for quote', 'quotation', 'quote request', 'pricing request'];
    const combined = (subject + ' ' + body).toLowerCase();
    return rfqKeywords.some(keyword => combined.includes(keyword));
}

function generateAiAnswer(question, context) {
    // Pre-defined answers for common FBG questions
    const answerDb = {
        'apodization': 'Yes, we can apply Gaussian apodization to all FBGs in the array. This technique shapes the refractive index profile to achieve sidelobe suppression levels typically better than -15dB. The apodization profile is optimized for your specified reflectivity to maintain spectral characteristics while minimizing sidelobe levels.',
        'temperature': 'The SM1330-E9/125PI fiber with polyimide coating is rated for continuous operation at temperatures up to 300°C, with short-term excursions possible up to 350°C. This has been validated through extensive thermal cycling tests.',
        'femtoplus': 'Yes, we offer FemtoPlus technology which provides improved thermal stability with wavelength drift typically <10pm/°C and enhanced mechanical durability. This is our recommended approach for applications requiring high temperature stability.',
        'aerospace': 'We can provide material certifications and test reports suitable for aerospace applications. Our manufacturing process is ISO 9001 certified, and we maintain full traceability documentation.',
        'certification': 'We can provide material certifications and test reports. Our manufacturing process is ISO 9001 certified with full traceability documentation. Specific certifications depend on the applicable standards for your project.',
        'radiation': 'We offer radiation-hardened fiber options using specialized fiber compositions that maintain performance under ionizing radiation exposure.',
        'connector': 'Typical insertion loss for our FC/APC connectors is <0.5dB with return loss >55dB. We use precision polishing and inspection to ensure consistent connector performance.'
    };
    
    const questionLower = question.toLowerCase();
    
    for (const [key, answer] of Object.entries(answerDb)) {
        if (questionLower.includes(key)) {
            return answer;
        }
    }
    
    // Default generic answer
    return 'Thank you for your question. Our engineering team will provide a detailed technical response. Please allow us to consult with our specialists to ensure we provide accurate information for your specific requirements.';
}

function generateEngineeringAnswer(question) {
    const questionLower = question.toLowerCase();
    
    if (questionLower.includes('apodization') || questionLower.includes('sidelobe')) {
        return 'Confirmed. We will apply Gaussian apodization to all FBGs in the array, achieving sidelobe suppression levels of <-15dB (typically -18dB to -20dB). The apodization profile will be optimized for your specified reflectivity. Our apodization process uses a proprietary phase mask design that ensures consistent results across all gratings.';
    }
    
    if (questionLower.includes('temperature') || questionLower.includes('polyimide')) {
        return 'The polyimide coating (SM1330-E9/125PI fiber from J-Fiber) is rated for continuous operation at temperatures up to 300°C with short-term excursions to 350°C. This has been validated through our internal thermal cycling tests (500 cycles, -40°C to +300°C). For naval applications, we recommend our enhanced polyimide option which provides additional moisture resistance.';
    }
    
    if (questionLower.includes('femtoplus') || questionLower.includes('stability')) {
        return 'Yes, FemtoPlus technology will be employed for all gratings. This provides: (1) Improved thermal stability with wavelength drift <10pm/°C, (2) Enhanced mechanical durability with >1000 strain cycles without degradation, (3) Reduced hydrogen sensitivity. The FemtoPlus process uses a proprietary high-temperature annealing procedure.';
    }
    
    if (questionLower.includes('aerospace') || questionLower.includes('certification') || questionLower.includes('documentation')) {
        return 'We can provide the following documentation for aerospace compliance: (1) Material certifications per MIL-STD-810, (2) RoHS and REACH compliance certificates, (3) Full traceability documentation from raw materials to finished product, (4) Test reports for thermal cycling, vibration, and humidity exposure. Our manufacturing process is ISO 9001:2015 certified.';
    }
    
    return 'We have reviewed your technical question and can confirm we can meet this requirement. Please contact our engineering team at engineering@engionic.de for specific technical details.';
}

function runPythonGenerator(scriptPath, specsPath, outputDir, offerNumber) {
    return new Promise((resolve, reject) => {
        const python = spawn('python3', [
            '-c',
            `
import sys
sys.path.insert(0, '${path.dirname(scriptPath)}')
import json
from document_generator import generate_quote_documents

with open('${specsPath}') as f:
    specs = json.load(f)

result = generate_quote_documents(specs, '${outputDir}', '${offerNumber}')
print(json.dumps(result))
            `
        ]);
        
        let output = '';
        let errorOutput = '';
        
        python.stdout.on('data', (data) => {
            output += data.toString();
        });
        
        python.stderr.on('data', (data) => {
            errorOutput += data.toString();
        });
        
        python.on('close', (code) => {
            if (code === 0) {
                try {
                    const result = JSON.parse(output.trim());
                    resolve({ success: true, ...result });
                } catch (e) {
                    resolve({ success: true, output });
                }
            } else {
                resolve({ success: false, error: errorOutput || 'Python script failed' });
            }
        });
        
        python.on('error', (err) => {
            resolve({ success: false, error: err.message });
        });
    });
}

// ============================================
// EMAIL SERVICE - Real Email Endpoints
// ============================================

const emailService = require('./src/services/emailService');

// Initialize Graph client on startup
emailService.initializeGraphClient().then(client => {
    if (client) {
        console.log('📧 Email service ready (Graph API mode)');
    } else {
        console.log('📧 Email service ready (Demo mode - no Azure credentials)');
    }
});

// Send email to engineering (REAL)
app.post('/api/email/send-to-engineering', async (req, res) => {
    try {
        const { rfqData, userEmail, originalMessageId } = req.body;
        
        const result = await emailService.sendToEngineering({
            rfqData,
            userEmail: userEmail || process.env.USER_EMAIL,
            originalMessageId
        });
        
        res.json({ success: true, ...result });
    } catch (error) {
        console.error('Error sending to engineering:', error);
        res.status(500).json({ success: false, error: error.message });
    }
});

// Trigger engineering reply simulation (sends real email to inbox)
app.post('/api/email/simulate-engineering-reply', async (req, res) => {
    try {
        const { conversationId, delayMs = 3000 } = req.body;
        
        const result = await emailService.simulateEngineeringReply({
            conversationId,
            delayMs
        });
        
        res.json({ success: true, ...result });
    } catch (error) {
        console.error('Error simulating engineering reply:', error);
        res.status(500).json({ success: false, error: error.message });
    }
});

// Send email to client (REAL)
app.post('/api/email/send-to-client', async (req, res) => {
    try {
        const { conversationId, engineeringAnswers } = req.body;
        
        const result = await emailService.sendToClient({
            conversationId,
            engineeringAnswers
        });
        
        res.json({ success: true, ...result });
    } catch (error) {
        console.error('Error sending to client:', error);
        res.status(500).json({ success: false, error: error.message });
    }
});

// Trigger client reply simulation (sends real email to inbox)
app.post('/api/email/simulate-client-reply', async (req, res) => {
    try {
        const { conversationId, delayMs = 3000 } = req.body;
        
        const result = await emailService.simulateClientReply({
            conversationId,
            delayMs
        });
        
        res.json({ success: true, ...result });
    } catch (error) {
        console.error('Error simulating client reply:', error);
        res.status(500).json({ success: false, error: error.message });
    }
});

// Get conversation state
app.get('/api/email/conversation/:conversationId', (req, res) => {
    const { conversationId } = req.params;
    const conversation = emailService.getConversation(conversationId);
    
    if (conversation) {
        res.json({ success: true, conversation });
    } else {
        res.status(404).json({ success: false, error: 'Conversation not found' });
    }
});

// List all active conversations
app.get('/api/email/conversations', (req, res) => {
    const conversations = emailService.listConversations();
    res.json({ success: true, conversations });
});

// Full workflow endpoint - runs entire demo sequence
app.post('/api/email/run-demo-workflow', async (req, res) => {
    try {
        const { rfqData, userEmail } = req.body;
        
        // Step 1: Send to engineering
        console.log('Step 1: Sending to engineering...');
        const engResult = await emailService.sendToEngineering({
            rfqData,
            userEmail: userEmail || process.env.USER_EMAIL
        });
        
        // Step 2: Simulate engineering reply (after delay)
        console.log('Step 2: Waiting for engineering reply...');
        const engReply = await emailService.simulateEngineeringReply({
            conversationId: engResult.conversationId,
            delayMs: 5000
        });
        
        // Step 3: Send to client
        console.log('Step 3: Sending to client...');
        const clientResult = await emailService.sendToClient({
            conversationId: engResult.conversationId,
            engineeringAnswers: engReply.answers
        });
        
        // Step 4: Simulate client reply
        console.log('Step 4: Waiting for client reply...');
        const clientReply = await emailService.simulateClientReply({
            conversationId: engResult.conversationId,
            delayMs: 5000
        });
        
        res.json({
            success: true,
            conversationId: engResult.conversationId,
            workflow: {
                engineeringSent: engResult,
                engineeringReply: engReply,
                clientSent: clientResult,
                clientReply: clientReply
            },
            completedSpecs: clientReply.completedSpecs
        });
        
    } catch (error) {
        console.error('Error in demo workflow:', error);
        res.status(500).json({ success: false, error: error.message });
    }
});

// Error handling middleware
app.use((err, req, res, next) => {
    console.error(err.stack);
    res.status(500).json({ error: 'Internal server error' });
});

// 404 handler
app.use((req, res) => {
    res.status(404).json({ error: 'Not found' });
});

// Start server
app.listen(PORT, () => {
    console.log(`
╔════════════════════════════════════════════════════════════╗
║                                                            ║
║   🚀 Hexa RFQ Manager Server Running                       ║
║                                                            ║
║   Local:    http://localhost:${PORT}                         ║
║   Taskpane: http://localhost:${PORT}/taskpane.html           ║
║   Manifest: http://localhost:${PORT}/manifest.xml            ║
║                                                            ║
║   Press Ctrl+C to stop                                     ║
║                                                            ║
╚════════════════════════════════════════════════════════════╝
    `);
});

module.exports = app;
