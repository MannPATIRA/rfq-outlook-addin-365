/**
 * Email content templates for the RFQ workflow (HTML format).
 */
const EMAIL_TEMPLATES = {
  engineeringReview: {
    subjectPrefix: 'Technical Review Required - RFQ #41260018 (NRL - 2 FBG Arrays)',
    body: `<div style="font-family: Arial, sans-serif; max-width: 600px; color: #111827;">
  <h2 style="color: #1e3a5f; border-bottom: 2px solid #1e3a5f; padding-bottom: 8px; margin: 0 0 16px 0;">Technical Review Required – RFQ #41260018</h2>
  <p style="margin: 0 0 16px 0;"><strong>Customer:</strong> NRL | <strong>Product:</strong> 2 FBG Arrays (qty 10)</p>
  <h3 style="color: #1e3a5f; font-size: 14px; margin: 16px 0 8px 0;">Technical questions requiring engineering input</h3>
  <ol style="margin: 0 0 16px 0; padding-left: 20px;">
    <li style="margin-bottom: 8px;"><strong>Apodization:</strong> Customer requests apodization for reduced sidelobe levels – confirm FemtoPlus compatibility and achievable SLSR.</li>
    <li style="margin-bottom: 8px;"><strong>Temperature range (-40°C to +85°C):</strong> Confirm annealing parameters for this operating range.</li>
    <li style="margin-bottom: 8px;"><strong>Polyimide coating:</strong> Confirm suitability for the specified temperature range.</li>
    <li style="margin-bottom: 8px;"><strong>Aerospace certification:</strong> Verify AS9100D availability if customer requests.</li>
  </ol>
  <h3 style="color: #1e3a5f; font-size: 14px; margin: 16px 0 8px 0;">Missing from customer RFQ (we will request)</h3>
  <ul style="margin: 0 0 16px 0; padding-left: 20px;">
    <li style="margin-bottom: 4px;">Reflectivity (%) – not specified</li>
    <li style="margin-bottom: 4px;">Calibration from/to temperatures – not specified</li>
    <li style="margin-bottom: 4px;">Delivery timeline – not specified</li>
  </ul>
  <p style="margin: 0; font-size: 13px; color: #4b5563;"><strong>Reference:</strong> NRL – 10x 2 FBG arrays, FC/APC both ends, SM1330-E9/125PI, 10050 mm total length, 50 mm FBG spacing, 1550.39 nm.</p>
</div>`,
  },

  engineeringReply: {
    body: `<div style="font-family: Arial, sans-serif; max-width: 600px; color: #111827;">
  <h2 style="color: #1e3a5f; border-bottom: 2px solid #1e3a5f; padding-bottom: 8px; margin: 0 0 16px 0;">Engineering Assessment – RFQ #41260018 (NRL – 2 FBG Arrays)</h2>
  <h3 style="color: #1e3a5f; font-size: 14px; margin: 16px 0 8px 0;">Technical answers</h3>
  <ol style="margin: 0 0 16px 0; padding-left: 20px;">
    <li style="margin-bottom: 10px;"><strong>Apodization:</strong> Confirmed. FemtoPlus technology with apodized gratings achieves minimum SLSR 8.0 dB.</li>
    <li style="margin-bottom: 10px;"><strong>Temperature range:</strong> Annealing at 300°C for 24 hours is suitable for -40°C to +85°C operating range.</li>
    <li style="margin-bottom: 10px;"><strong>Polyimide coating:</strong> Suitable for continuous use up to +300°C. No issues for specified range.</li>
    <li style="margin-bottom: 10px;"><strong>Aerospace certification:</strong> AS9100D available upon request. No additional lead time.</li>
  </ol>
  <h3 style="color: #1e3a5f; font-size: 14px; margin: 16px 0 8px 0;">Clarifications to request from customer</h3>
  <ul style="margin: 0 0 16px 0; padding-left: 20px;">
    <li style="margin-bottom: 4px;"><strong>Confirm:</strong> Is 8.0 dB SLSR acceptable as minimum (RFQ says "nominal")?</li>
    <li style="margin-bottom: 4px;"><strong>Confirm:</strong> 5-point or 10-point calibration data required?</li>
  </ul>
  <p style="margin: 0; padding: 10px; background: #e8f0f8; border-left: 4px solid #1e3a5f;">Ready to send clarification email requesting: reflectivity %, calibration range, delivery timeline, plus above confirmations.</p>
</div>`,
  },

  customerClarification: {
    subjectPrefix: 'RE: RFQ for 2 FBG Arrays - Clarification Required',
    body: `<div style="font-family: Arial, sans-serif; max-width: 600px; color: #111827;">
  <p style="margin: 0 0 16px 0;">Thank you for your RFQ (NRL – 2 FBG arrays, Offer #41260018). We have reviewed your requirements and have the following.</p>
  <h3 style="color: #1e3a5f; font-size: 14px; margin: 16px 0 8px 0;">Answers to your questions</h3>
  <ol style="margin: 0 0 16px 0; padding-left: 20px;">
    <li style="margin-bottom: 8px;"><strong>Apodization:</strong> Yes, we provide apodized FBGs as standard. Our FemtoPlus technology achieves a minimum SLSR of 8 dB.</li>
    <li style="margin-bottom: 8px;"><strong>Maximum operating temperature (polyimide):</strong> Suitable for continuous operation from -40°C to +300°C; your range (-40°C to +85°C) is fully supported.</li>
    <li style="margin-bottom: 8px;"><strong>FemtoPlus technology:</strong> Yes, it is used for improved thermal stability and long-term reliability.</li>
    <li style="margin-bottom: 8px;"><strong>Annealing:</strong> We use annealing at 300°C for 24 hours for your operating range.</li>
    <li style="margin-bottom: 8px;"><strong>Aerospace certification:</strong> AS9100D certification is available upon request.</li>
  </ol>
  <h3 style="color: #1e3a5f; font-size: 14px; margin: 16px 0 8px 0;">Information needed to prepare quote</h3>
  <ul style="margin: 0 0 16px 0; padding-left: 20px;">
    <li style="margin-bottom: 4px;"><strong>Reflectivity (%):</strong> Please specify required reflectivity (typical range 10–90%).</li>
    <li style="margin-bottom: 4px;"><strong>Calibration range:</strong> Please confirm the from/to temperatures for calibration data.</li>
    <li style="margin-bottom: 4px;"><strong>Delivery:</strong> Please indicate required delivery date or acceptable lead time.</li>
  </ul>
  <h3 style="color: #1e3a5f; font-size: 14px; margin: 16px 0 8px 0;">Please confirm</h3>
  <ul style="margin: 0 0 16px 0; padding-left: 20px;">
    <li style="margin-bottom: 4px;">Is 8.0 dB SLSR acceptable as a minimum? Your RFQ lists "nominal" – please clarify if this is a minimum acceptance criterion.</li>
    <li style="margin-bottom: 4px;">For calibration, do you require 5-point temperature data (standard) or 10-point higher-resolution characterisation?</li>
  </ul>
  <p style="margin: 0;">We look forward to your reply.</p>
</div>`,
  },

  customerReplyWithDetails: {
    body: `<div style="font-family: Arial, sans-serif; max-width: 600px; color: #111827;">
  <p style="margin: 0 0 16px 0;">Thank you for the clarifications. Please find our responses below:</p>
  <h3 style="color: #1e3a5f; font-size: 14px; margin: 16px 0 8px 0;">Missing details</h3>
  <ul style="margin: 0 0 16px 0; padding-left: 20px;">
    <li style="margin-bottom: 6px;"><strong>Reflectivity:</strong> 10% (as per standard for this sensing application).</li>
    <li style="margin-bottom: 6px;"><strong>Calibration range:</strong> -40°C to +85°C (match operating range).</li>
    <li style="margin-bottom: 6px;"><strong>Delivery:</strong> Lead time of 8 weeks from order confirmation is acceptable.</li>
  </ul>
  <h3 style="color: #1e3a5f; font-size: 14px; margin: 16px 0 8px 0;">Confirmations</h3>
  <ul style="margin: 0 0 16px 0; padding-left: 20px;">
    <li style="margin-bottom: 6px;"><strong>SLSR:</strong> Yes, 8.0 dB minimum is acceptable as the acceptance criterion.</li>
    <li style="margin-bottom: 6px;"><strong>Calibration data:</strong> Standard 5-point characterisation is sufficient.</li>
  </ul>
  <p style="margin: 0 0 16px 0;">We confirm the remaining specification as per our RFQ (2 FBGs, 10050 mm total length, 50 mm spacing, 1550.39 nm, SM1330-E9/125PI, FC/APC both ends, quantity 10). Please send the formal quote and specification sheet at your earliest convenience.</p>
</div>`,
  },

  finalQuoteToCustomer: {
    subjectPrefix: 'RE: RFQ for 2 FBG Arrays - Quote #41260018',
    body: `<div style="font-family: Arial, sans-serif; max-width: 600px; color: #111827;">
  <p style="margin: 0 0 16px 0;">Dear Customer,</p>
  <p style="margin: 0 0 16px 0;">Thank you for your confirmation. Please find attached:</p>
  <ul style="margin: 0 0 16px 0; padding-left: 20px;">
    <li style="margin-bottom: 4px;">Specification Sheet – 41260018 NRL.xlsx</li>
    <li style="margin-bottom: 4px;">Quote Document – 41260018 NRL.pdf</li>
  </ul>
  <h3 style="color: #1e3a5f; font-size: 14px; margin: 16px 0 8px 0;">Summary</h3>
  <ul style="margin: 0 0 16px 0; padding-left: 20px;">
    <li style="margin-bottom: 4px;">10x 2 FBG arrays with FC/APC connectors on both ends (see spec sheet 41260018).</li>
    <li style="margin-bottom: 4px;"><strong>Net amount:</strong> €2,206.40 (tax free, NON-EU).</li>
    <li style="margin-bottom: 4px;"><strong>Terms:</strong> Payment in advance, EXW (Incoterms 2020).</li>
    <li style="margin-bottom: 4px;">Please return signed specification sheet with order.</li>
    <li style="margin-bottom: 4px;">This offer is valid for 30 days.</li>
  </ul>
  <p style="margin: 0;">If you have any further questions, please do not hesitate to contact us.</p>
  <p style="margin: 16px 0 0 0;">Best regards</p>
</div>`,
  },
};

if (typeof window !== 'undefined') {
  window.EMAIL_TEMPLATES = EMAIL_TEMPLATES;
}
