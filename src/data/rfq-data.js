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
    { id: 'm1', field: 'Reflectivity (%)', importance: 'critical', note: 'Nominal reflectivity not specified; only tolerances given.' },
    { id: 'm2', field: 'FWHM tolerance (nm)', importance: 'critical', note: 'Exact FWHM tolerance requirement to be confirmed.' },
    { id: 'm3', field: 'SLSR minimum requirement (dB)', importance: 'critical', note: 'Sheet shows 8.0 dB nominal; confirm if this is the minimum required.' },
    { id: 'm4', field: 'Calibration temperature range', importance: 'high', note: 'Calibration from/to temperatures not specified.' },
    { id: 'm5', field: 'Delivery timeline preference', importance: 'high', note: 'Required delivery date or lead time not specified.' },
  ],

  questionsToAsk: [
    'What reflectivity percentage do you require for the FBGs (typical range 10–99%)?',
    'What is your acceptable FWHM tolerance (we offer nominal 0.09 nm ± 0.02 nm)?',
    'Can you confirm the minimum SLSR requirement (we quote 8.0 dB nominal)?',
    'What calibration temperature range do you need (from/to °C) for temperature sensing?',
    'What is your required delivery date or lead time for the 10 units?',
  ],
};

if (typeof window !== 'undefined') {
  window.RFQ_DATA = RFQ_DATA;
}
