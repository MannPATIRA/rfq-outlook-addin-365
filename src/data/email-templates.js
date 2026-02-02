/**
 * Email content templates for the RFQ workflow.
 * Used when sending to engineering, replying to customer, or sending final quote.
 */
const EMAIL_TEMPLATES = {
  /** Subject and body for the email sent from user to engineering team (Notify Engineering). */
  engineeringReview: {
    subjectPrefix: 'Technical Review Required - RFQ #41260018 (NRL - 2 FBG Arrays)',
    body: `Key concerns requiring engineering input:

1. Customer requests apodization for reduced sidelobe levels – please confirm FemtoPlus compatibility and SLSR capability.
2. Operating temperature range -40°C to +85°C – confirm annealing parameters and calibration range.
3. Aerospace certification requirements – verify AS9100D or other compliance and availability.
4. Polyimide coating – confirm max operating temperature and long-term stability for this range.
5. Reflectivity, FWHM tolerance, and delivery timeline – customer has not yet specified; we will request. Please confirm our standard offerings (e.g. nominal 0.09 nm FWHM ±0.02 nm, 8 dB SLSR).

Reference: NRL – 10x 2 FBG arrays, FC/APC both ends, SM1330-E9/125PI, 10050 mm total length, 50 mm FBG spacing, 1550.39 nm.`,
  },

  /** Automated reply from engineering-team back to the user. */
  engineeringReply: {
    body: `Engineering Assessment – RFQ #41260018 (NRL – 2 FBG Arrays):

1. Apodization: Confirmed. FemtoPlus technology with apodized gratings; minimum SLSR 8.0 dB. Standard offering.
2. Temperature range (-40°C to +85°C): Annealing at 300°C for 24 hours recommended. Calibration from/to range can be provided upon customer confirmation.
3. Aerospace certification: AS9100D certification available upon request. No additional lead time for standard documentation.
4. Polyimide coating: Suitable for continuous use up to +300°C. No issues for -40°C to +85°C.
5. Standards: Reflectivity nominal as per spec sheet (confirm with customer if % required). FWHM nominal 0.09 nm ±0.02 nm. SLSR 8.0 dB. Delivery TBD once customer confirms timeline.

Ready for clarification email to customer with above answers and request for missing details (reflectivity %, calibration range, delivery date).`,
  },

  /** Reply from user to customer-1 requesting clarifications (after engineering review or direct). */
  customerClarification: {
    subjectPrefix: 'RE: RFQ for 2 FBG Arrays - Clarification Required',
    body: `Thank you for your RFQ (NRL – 2 FBG arrays, Offer #41260018). We have reviewed your requirements and have the following.

ANSWERS TO YOUR QUESTIONS:

1. Apodization: Yes, we provide apodized FBGs as standard. Our FemtoPlus technology achieves a minimum SLSR of 8 dB.
2. Maximum operating temperature (polyimide-coated): Suitable for continuous operation from -40°C to +300°C; your range (-40°C to +85°C) is fully supported.
3. FemtoPlus technology: Yes, it is used for improved thermal stability and long-term reliability.
4. Annealing for your operating range: We use annealing at 300°C for 24 hours. Calibration over your temperature range can be provided – please specify from/to temperatures.
5. Aerospace certification: AS9100D certification is available upon request.

ADDITIONAL INFORMATION NEEDED:

To prepare an accurate quote, please confirm or specify:
- Required reflectivity percentage for the FBGs (typical range 10–99%).
- Acceptable FWHM tolerance (we offer nominal 0.09 nm ± 0.02 nm).
- Confirmation of minimum SLSR requirement (we quote 8.0 dB nominal).
- Calibration temperature range (from/to °C) if applicable.
- Required delivery date or lead time for the 10 units.

We look forward to your reply.`,
  },

  /** Automated reply from customer-1 back to user with “missing details” filled in. */
  customerReplyWithDetails: {
    body: `Thank you for the clarifications. Please find our responses below:

- Reflectivity: 10% (as per standard for this application).
- FWHM tolerance: 0.09 nm ± 0.02 nm is acceptable.
- SLSR: 8.0 dB minimum confirmed.
- Calibration temperature range: -40°C to +85°C (match operating range).
- Delivery: We require delivery by [date] / lead time of 8 weeks from order confirmation.

We confirm the remaining specification as per our RFQ (2 FBGs, 10050 mm total length, 50 mm spacing, 1550.39 nm, SM1330-E9/125PI, FC/APC both ends, quantity 10). Please send the formal quote and specification sheet at your earliest convenience.`,
  },

  /** Final reply from user to customer-1 with quote and attachments. */
  finalQuoteToCustomer: {
    subjectPrefix: 'RE: RFQ for 2 FBG Arrays - Quote #41260018',
    body: `Dear Customer,

Thank you for your confirmation. Please find attached:

1. Specification Sheet – 41260018 NRL.xlsx  
2. Quote Document – 41260018 NRL.pdf  

Summary:
- 10x 2 FBG arrays with FC/APC connectors on both ends (see spec sheet 41260018).
- Net amount: €2,206.40 (tax free, NON-EU).
- Terms: Payment in advance, EXW (Incoterms 2020).
- Please return signed specification sheet with order.
- This offer is valid for 30 days.

If you have any further questions, please do not hesitate to contact us.

Best regards`,
  },
};

if (typeof window !== 'undefined') {
  window.EMAIL_TEMPLATES = EMAIL_TEMPLATES;
}
