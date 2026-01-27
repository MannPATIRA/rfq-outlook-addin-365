# Hexa RFQ Manager - Outlook Add-in

<p align="center">
  <img src="assets/logo.png" alt="Hexa RFQ Manager" width="200">
</p>

**Technical Sales RFQ Management System for Fiber Bragg Grating (FBG) Products**

This Outlook add-in automates the complete Request for Quote (RFQ) workflow for optical fiber products, specifically Fiber Bragg Grating (FBG) arrays.

---

## 🎯 Features

### Core Functionality
- **Automatic RFQ Analysis**: Extracts technical specifications from incoming RFQ emails
- **Missing Spec Detection**: Identifies required specifications not provided in the RFQ
- **Client Question Parsing**: Extracts and categorizes technical questions from clients
- **AI-Powered Answers**: Generates suggested responses for common technical questions
- **Engineering Workflow**: Routes complex questions to engineering team with context
- **Document Generation**: Creates professional Excel spec sheets and PDF quotes
- **Email Threading**: Maintains proper email threading throughout the workflow

### Workflow Steps
1. **RFQ Receipt** → Automatic extraction of specs, questions, and missing info
2. **Engineering Consultation** → Send technical questions to engineering team
3. **Client Clarification** → Reply to client with answers and spec requests
4. **Spec Completion** → Process client response with complete specifications
5. **Quote Generation** → Generate Excel spec sheet and PDF quote document
6. **Final Delivery** → Send quote to client with attached documents

---

## 📋 Prerequisites

- **Node.js** 18.0 or higher
- **Python** 3.8 or higher
- **Microsoft Outlook** (Desktop or Web)
- **Microsoft 365** account (for add-in sideloading)

### Python Dependencies
```bash
pip install openpyxl reportlab
```

### Node.js Dependencies
```bash
npm install
```

---

## 🚀 Quick Start

### 1. Clone and Install

```bash
# Clone the repository
git clone https://github.com/hexa-ai/rfq-outlook-addin.git
cd rfq-outlook-addin

# Install Node.js dependencies
npm install

# Install Python dependencies
pip install openpyxl reportlab
```

### 2. Configure Environment

```bash
# Copy environment template
cp .env.example .env

# Edit .env with your settings
nano .env
```

### 3. Start the Server

```bash
# Development mode with auto-reload
npm run dev

# Or production mode
npm start
```

Server will start at `https://localhost:3000`

### 4. Install Add-in in Outlook

#### Option A: Sideload in Outlook Desktop (Windows)

1. Open Outlook
2. Go to **File** → **Manage Add-ins** or **Options** → **Add-ins**
3. Click **My Add-ins** → **Add a custom add-in** → **Add from file**
4. Select `manifest.xml` from the project directory

#### Option B: Sideload in Outlook Web

1. Go to [Outlook Web](https://outlook.office.com)
2. Click the **gear icon** → **Manage add-ins**
3. Click **My add-ins** → **Add a custom add-in** → **Add from URL**
4. Enter: `https://localhost:3000/manifest.xml`

#### Option C: Centralized Deployment (Admin)

1. Go to [Microsoft 365 Admin Center](https://admin.microsoft.com)
2. Navigate to **Settings** → **Integrated apps**
3. Click **Upload custom apps**
4. Upload `manifest.xml`

---

## 📁 Project Structure

```
rfq-outlook-addin/
├── manifest.xml              # Outlook add-in manifest
├── package.json              # Node.js dependencies
├── server.js                 # Express server
├── .env                      # Environment configuration
│
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html    # Add-in UI
│   │   └── taskpane.js      # Add-in logic
│   │
│   ├── commands/
│   │   └── commands.js      # Ribbon commands
│   │
│   └── helpers/
│       ├── document_generator.py  # Excel/PDF generation
│       └── demo_automation.js     # Demo email automation
│
├── assets/
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-64.png
│   ├── icon-80.png
│   └── icon-128.png
│
├── demo-data/
│   └── sample_rfq.eml       # Sample RFQ email
│
└── output/                  # Generated documents
```

---

## 🎮 Demo Mode

The add-in includes a full demo mode that simulates the complete workflow without requiring real email setup.

### Running the Demo

1. Start the server: `npm start`
2. Open `https://localhost:3000/taskpane.html` in a browser
3. The demo will automatically run through the workflow

### Demo Features
- Simulated RFQ extraction
- Pre-populated technical questions
- AI-generated answers
- Simulated engineering responses
- Complete specification processing
- Document generation

---

## 📧 Email Configuration (Production)

For production use, configure email addresses for automated responses:

```env
# .env file

# Engineering team email (for technical questions)
ENGINEERING_EMAIL=engineering@yourcompany.com

# Client email (RFQ sender - for demo only)
CLIENT_EMAIL=client@demo.yourcompany.com

# Salesperson email (where add-in runs)
SALES_EMAIL=sales@yourcompany.com

# SMTP configuration (for automated responses)
SMTP_HOST=smtp.yourprovider.com
SMTP_PORT=587
SMTP_USER=your-email@yourcompany.com
SMTP_PASS=your-app-password
```

---

## 📊 Document Generation

### Excel Specification Sheet

The add-in generates a multi-sheet Excel workbook matching the professional format:

| Sheet | Contents |
|-------|----------|
| Sensor Specification | Main specs, FBG parameters, wavelength table |
| Fiber Specification | Fiber details, linked to sensor sheet |
| Definitions | Lookup tables for fiber types, contacts |
| Drawings | Configuration references |

### PDF Quote Document

Professional quote PDF including:
- Company header and branding
- Customer information
- Itemized pricing table
- Terms and conditions
- Dual-use classification notice
- Bank details and legal info

---

## 🔧 API Reference

### Extract Specifications
```http
POST /api/extract-specs
Content-Type: application/json

{
  "emailBody": "RFQ email content...",
  "emailSubject": "RFQ: FBG Arrays for Project"
}
```

### Generate Documents
```http
POST /api/generate-documents
Content-Type: application/json

{
  "specifications": {
    "Customer": "NRL",
    "Quantity": "10 pcs",
    ...
  },
  "offerNumber": "41260018"
}
```

### Simulate Engineering Response
```http
POST /api/simulate-engineering-response
Content-Type: application/json

{
  "questions": [
    { "id": 1, "question": "Can you provide apodization?" }
  ]
}
```

---

## 🛠️ Customization

### Adding New Fiber Types

Edit `src/helpers/document_generator.py`:

```python
fiber_types = [
    ('Your-New-Fiber', 'Core Type', 10.0, 125, 'Coating', 'Manufacturer', 'ITU-T'),
    ...
]
```

### Customizing AI Answers

Edit `server.js` `generateAiAnswer()` function:

```javascript
const answerDb = {
    'your_keyword': 'Your custom answer for this topic...',
    ...
};
```

### Modifying Quote Template

Edit `src/helpers/document_generator.py` class `QuoteDocumentGenerator`:

- `generate_excel_spec_sheet()` - Excel structure
- `generate_pdf_quote()` - PDF layout

---

## 🔒 Security Notes

- The add-in requests `ReadWriteMailbox` permissions
- All data processing happens locally (no external API calls)
- Document generation is server-side only
- For production: Use HTTPS and proper authentication

---

## 🐛 Troubleshooting

### Add-in not loading
- Check that the server is running on `https://localhost:3000`
- Verify SSL certificate is trusted
- Clear Outlook cache and reload

### Document generation fails
- Ensure Python dependencies are installed
- Check write permissions for `output/` directory
- Verify Python is in system PATH

### Email threading not working
- Message IDs must be preserved across the workflow
- Check that `originalRfqMessageId` is being stored

---

## 📄 License

MIT License - see [LICENSE](LICENSE) file

---

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Open a Pull Request

---

## 📞 Support

- **Documentation**: [docs.hexa.ai](https://docs.hexa.ai)
- **Issues**: [GitHub Issues](https://github.com/hexa-ai/rfq-outlook-addin/issues)
- **Email**: support@hexa.ai

---

<p align="center">
  Built with ❤️ by <a href="https://hexa.ai">Hexa AI</a>
</p>
