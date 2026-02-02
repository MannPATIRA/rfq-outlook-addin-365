/**
 * Preset RFQ data for the NRL 2 FBG array request (Offer #41260018).
 * Used when the add-in is opened on the initial RFQ email from customer-1.
 */
const RFQ_DATA = {
  technicalSpecs: {
    quantity: 10,
    offerNumber: '41260018',
    customer: 'NRL',
    fiberType: 'SM1330-E9/125PI',
    fiberCoreType: 'Single Mode',
    claddingDiameter: '125 μm',
    coatingMaterial: 'Polyimide',
    fiberManufacturer: 'J-Fiber',
    fiberProvidedBy: 'eFG',
    totalFiberLength: '10050 mm',
    numberOfFBGs: 2,
    connectorType: 'FC/APC both ends',
    wavelengthNm: '1550.39',
    fwhmNominalNm: '0.09',
    fwhmTolerancePlus: '0.02',
    fwhmToleranceMinus: '0.02',
    minSlsrDb: '8.0',
    fbgLengthMm: '12.0',
    fbgLengthTolerancePlus: '2.0',
    fbgLengthToleranceMinus: '2.0',
    fbgSpacingMm: '50',
    toleranceFbgSpacingPercent: '1',
    label: 'on spool',
    remarks: 'Kein Faserwechsel ohne Absprache; Toleranz erstes bis letztes FBG max 5mm',
    apodized: 'yes',
    femtoPlus: '0.00',
  },

  customerQuestions: [
    {
      id: 'q1',
      question: 'Can you provide apodization for reduced sidelobe levels?',
      answer: 'Yes, we offer apodized FBGs as standard. Our FemtoPlus technology uses apodization to achieve a minimum SLSR of 8 dB, reducing sidelobes and improving sensor performance in dense wavelength division multiplexing applications.',
    },
    {
      id: 'q2',
      question: 'What is the maximum operating temperature for polyimide-coated fibers?',
      answer: 'Polyimide-coated fibers are suitable for continuous operation from -40°C to +300°C. For your specified operating range (-40°C to +85°C), no special annealing beyond our standard process is required. We can provide annealing and calibration details for your exact range upon order.',
    },
    {
      id: 'q3',
      question: 'Will FemtoPlus technology be used for improved thermal stability?',
      answer: 'Yes. FemtoPlus technology is applied for improved thermal stability and long-term reliability. Our standard process includes annealing at 300°C for 24 hours to stabilize the gratings for your operating range.',
    },
    {
      id: 'q4',
      question: 'What annealing temperatures do you use for our operating range (-40°C to +85°C)?',
      answer: 'For the -40°C to +85°C operating range, we use annealing at 300°C for 24 hours. This stabilizes the FBG response and minimizes drift. Calibration can be performed over your specified temperature range; please confirm your desired calibration from/to temperatures.',
    },
    {
      id: 'q5',
      question: 'Can you provide certification for aerospace applications?',
      answer: 'We can provide AS9100D certification upon request. Our FBG arrays meet the requirements for aerospace and defense applications. Please specify if you need additional documentation (e.g., material certifications, test reports) for your procurement process.',
    },
  ],

  missingDetails: [
    { id: 'm1', field: 'Reflectivity (%)', importance: 'critical', note: 'Required reflectivity not specified – needed to set FBG writing parameters (typical 10–90%).' },
    { id: 'm2', field: 'Calibration from/to (°C)', importance: 'critical', note: 'Customer states operating range but not calibration data points required.' },
    { id: 'm3', field: 'Delivery lead time', importance: 'high', note: 'No delivery date or lead time given – affects production scheduling.' },
  ],

  questionsToAsk: [
    'Please confirm that 8.0 dB SLSR is an acceptable minimum – your RFQ lists "nominal" which may imply a different acceptance criterion.',
    'For the calibration certificate, do you require 5-point temperature data (standard) or higher-resolution 10-point characterisation?',
  ],
};


if (typeof window !== 'undefined') {
  window.RFQ_DATA = RFQ_DATA;
}
