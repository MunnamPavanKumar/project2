const express = require('express');
const router = express.Router();
const multer = require('multer');
const xlsx = require('exceljs');
const db = require('../db/connection');
const fs = require('fs');
const path = require('path');
const archiver = require('archiver');
const authenticateToken = require('../middleware/authMiddleware');
const upload = multer({ dest: 'uploads/' });



let processedData = new Map();

const offices = [
  { code: 4000, name: "Southern Regional Office", region: "Southern Region", lineId: 10 },
  { code: 4000, name: "Southern Regional Office", region: "Southern Region", lineId: 20},
  { code: 4001, name: "Arakkonam AFS", region: "Tamil Nadu State", lineId: 10 },
  { code: 4002, name: "Coimbatore AFS", region: "Tamil Nadu State", lineId: 20},
  { code: 4003, name: "Meenambakkam AFS", region: "Tamil Nadu State", lineId: 30},
  { code: 4004, name: "Trichy AFS", region: "Tamil Nadu State", lineId: 40 },
  { code: 4005, name: "Sulur AFS", region: "Tamil Nadu State", lineId: 50 },
  { code: 4006, name: "Madurai AFS", region: "Tamil Nadu State", lineId: 60 },
  { code: 4007, name: "Tambaram AFS", region: "Tamil Nadu State", lineId: 70},
  { code: 4028, name: "Tuticorin AFS", region: "Tamil Nadu State", lineId: 80 },
  { code: 4038, name: "Ramnad AFS", region: "Tamil Nadu State", lineId: 90},
  { code: 4076, name: "Chennai Drum Plant", region: "Tamil Nadu State", lineId: 100},
  { code: 4100, name: "Tamil Nadu State Office", region: "Tamil Nadu State", lineId: 110},
  { code: 4101, name: "Coimbatore DO", region: "Tamil Nadu State", lineId: 111},
  { code: 4102, name: "Chennai DO", region: "Tamil Nadu State", lineId: 112},
  { code: 4103, name: "Madurai DO", region: "Tamil Nadu State", lineId: 113},
  { code: 4104, name: "Salem Divisional Office", region: "Tamil Nadu State", lineId: 114},
  { code: 4105, name: "Trichy DO", region: "Tamil Nadu State", lineId: 115},
  { code: 4111, name: "Coimbatore Indane DO", region: "Tamil Nadu State", lineId: 116},
  { code: 4112, name: "Chennai Indane DO", region: "Tamil Nadu State", lineId: 117},
  { code: 4113, name: "Madurai Indane DO", region: "Tamil Nadu State", lineId: 118},
  { code: 4114, name: "Trichy Indane DO", region: "Tamil Nadu State", lineId: 119},
  { code: 4121, name: "CPCL Manali", region: "Tamil Nadu State", lineId: 120},
  { code: 4125, name: "Chennai Terminal Foreshore", region: "Tamil Nadu State", lineId: 154},
  { code: 4126, name: "Madurai Terminal", region: "Tamil Nadu State", lineId: 121},
  { code: 4127, name: "Chennai Terminal - Korukkupet", region: "Tamil Nadu State", lineId: 122},
  { code: 4129, name: "Chennai Terminal - Tondiarpet", region: "Tamil Nadu State", lineId: 123},
  { code: 4130, name: "Tuticorin Terminal", region: "Tamil Nadu State", lineId: 124},
  { code: 4133, name: "Trichy Terminal", region: "Tamil Nadu State", lineId: 125},
  { code: 4134, name: "Chennai FST", region: "Tamil Nadu State", lineId: 126},
  //{ code: 4135, name: "Narimanam Plant", region: "Tamil Nadu State", lineId: },
  { code: 4136, name: "LBP Chennai", region: "Tamil Nadu State", lineId: 128},
  { code: 4141, name: "Asanur Terminal", region: "Tamil Nadu State", lineId: 129},
  { code: 4149, name: "Coimbatore Terminal", region: "Tamil Nadu State", lineId: 130},
  { code: 4150, name: "Salem Terminal", region: "Tamil Nadu State", lineId: 164},
  { code: 4150, name: "Salem Terminal", region: "Tamil Nadu State", lineId: 184},
  { code: 4159, name: "IOCL CO HPC Ennore", region: "Tamil Nadu State", lineId: 131},
  { code: 4167, name: "IOC Cell RIL Ennore", region: "Tamil Nadu State", lineId: 132},
  { code: 4170, name: "Chennai POL Jetty", region: "Tamil Nadu State", lineId: 133},
  { code: 4171, name: "LPG BP Ennore", region: "Tamil Nadu State", lineId: 134},
  { code: 4172, name: "LPG BP Salem", region: "Tamil Nadu State", lineId: 135},
  { code: 4174, name: "Pondichery", region: "Pondichery", lineId: 10},
  { code: 4175, name: "LPG BP Madurai", region: "Tamil Nadu State", lineId: 136},
  { code: 4176, name: "LPG BP Mayiladuthurai", region: "Tamil Nadu State", lineId: 137},
  { code: 4177, name: "LPG BP Erode", region: "Tamil Nadu State", lineId: 138},
  { code: 4179, name: "LPG BP Coimbatore", region: "Tamil Nadu State", lineId: 174},
  { code: 4181, name: "LPG BP Trichy", region: "Tamil Nadu State", lineId: 139},
  { code: 4183, name: "LPG BP Chengelpet", region: "Tamil Nadu State", lineId: 140},
  { code: 4184, name: "LPG BP Mannargudi", region: "Tamil Nadu State", lineId: 141},
  { code: 4185, name: "Coimbatore Bottling Plant", region: "Tamil Nadu State", lineId: 142},
  { code: 4187, name: "LPG BP Ilayangudi", region: "Tamil Nadu State", lineId: 143},
  { code: 4188, name: "LPG BP Tirunelveli", region: "Tamil Nadu State", lineId: 144}
];
function calculateConsumedQuantityFromPreviousQuarters(amcFrom, amcTo, currentQuarter, noOfItems) {
  const quarterRanges = {
    'Q1': { start: new Date('2024-06-01'), end: new Date('2024-08-31') },
    'Q2': { start: new Date('2024-09-01'), end: new Date('2024-11-30') },
    'Q3': { start: new Date('2024-12-01'), end: new Date('2025-02-28') },
    'Q4': { start: new Date('2025-03-01'), end: new Date('2025-05-31') },
    'Q5': { start: new Date('2025-06-01'), end: new Date('2025-08-31') },
    'Q6': { start: new Date('2025-09-01'), end: new Date('2025-11-30') },
    'Q7': { start: new Date('2025-12-01'), end: new Date('2026-02-28') },
    'Q8': { start: new Date('2026-03-01'), end: new Date('2026-05-31') },
    'Q9': { start: new Date('2026-06-01'), end: new Date('2026-08-31') },
    'Q10': { start: new Date('2026-09-01'), end: new Date('2026-11-30') },
    'Q11': { start: new Date('2026-12-01'), end: new Date('2027-02-28') },
    'Q12': { start: new Date('2027-03-01'), end: new Date('2027-05-31') }
  };

  const currentQuarterIndex = parseInt(currentQuarter.replace('Q', ''));
  let totalConsumedQuantity = 0;
  let quarterBreakdown = [];

  // Calculate consumed quantity for all previous quarters
  for (let i = 1; i < currentQuarterIndex; i++) {
    const quarterKey = `Q${i}`;
    const quarterRange = quarterRanges[quarterKey];
    
    if (!quarterRange) continue;

    const serviceDays = calculateServiceDays(amcFrom, amcTo, quarterRange.start, quarterRange.end);
    
    // Only add if service days > 0 (within AMC range)
    if (serviceDays > 0) {
      const quarterConsumption = noOfItems * serviceDays;
      totalConsumedQuantity += quarterConsumption;
      quarterBreakdown.push({
        quarter: quarterKey,
        serviceDays: serviceDays,
        consumption: quarterConsumption
      });
    }
  }

  return {
    totalConsumedQuantity,
    quarterBreakdown
  };
}



function getQuarterRange(quarter) {
  // Define base year - you can modify this as needed
  const baseYear = 2024;
  
  const ranges = {
    Q1: [`${baseYear}-06-01`, `${baseYear}-08-31`],        // Jun-Aug 2024
    Q2: [`${baseYear}-09-01`, `${baseYear}-11-30`],        // Sep-Nov 2024
    Q3: [`${baseYear}-12-01`, `${baseYear + 1}-02-28`],    // Dec 2024 - Feb 2025
    Q4: [`${baseYear + 1}-03-01`, `${baseYear + 1}-05-31`], // Mar-May 2025
    Q5: [`${baseYear + 1}-06-01`, `${baseYear + 1}-08-31`], // Jun-Aug 2025
    Q6: [`${baseYear + 1}-09-01`, `${baseYear + 1}-11-30`], // Sep-Nov 2025
    Q7: [`${baseYear + 1}-12-01`, `${baseYear + 2}-02-28`], // Dec 2025 - Feb 2026
    Q8: [`${baseYear + 2}-03-01`, `${baseYear + 2}-05-31`], // Mar-May 2026
    Q9: [`${baseYear + 2}-06-01`, `${baseYear + 2}-08-31`], // Jun-Aug 2026
    Q10: [`${baseYear + 2}-09-01`, `${baseYear + 2}-11-30`], // Sep-Nov 2026
    Q11: [`${baseYear + 2}-12-01`, `${baseYear + 3}-02-28`], // Dec 2026 - Feb 2027
    Q12: [`${baseYear + 3}-03-01`, `${baseYear + 3}-05-31`], // Mar-May 2027
};
  
  return ranges[quarter] || ranges['Q1'];; // Default to Q1 if quarter not found
}


function calculateServiceDays(amcFrom, amcTo, qStart, qEnd) {
  const from = new Date(amcFrom);
  const to = new Date(amcTo);
  const start = new Date(qStart);
  const end = new Date(qEnd);

  if (to < start || from > end) return 0;
  const actualStart = from > start ? from : start;
  const actualEnd = to < end ? to : end;

  const timeDiff = actualEnd.getTime() - actualStart.getTime();
  return Math.floor(timeDiff / (1000 * 60 * 60 * 24)) + 1;
}

function generatePONumber() {
  return '70134' + Math.floor(Math.random() * 1000).toString().padStart(3, '0');
}

function getTaxCode(locationCode) {
  const gqCodes = [4000, 4005, 4100, 4101, 4102, 4103, 4104, 4105, 4111, 4112, 4113, 4114];
  const grCodes = [4076];
  
  if (gqCodes.includes(locationCode)) return 'GQ';
  if (grCodes.includes(locationCode)) return 'GR';
  return 'GP'; // Default for most locations
}

// Helper function to sanitize folder names
function sanitizeFolderName(name) {
  return name.replace(/[<>:"/\\|?*]/g, '_').replace(/\s+/g, '_');
}

// Store generated report filenames to track them
let generatedReports = {
  locationWise: null,
  regionWise: null,
  quarter: null,
  yearRange: null
};

async function generateLocationWiseReport(quarter) {
  try {
    const workbook = new xlsx.Workbook();
    const worksheet = workbook.addWorksheet('Location Wise Report');

    // Define headers
    const headers = ['S.No', 'Line Item', 'Location Code', 'Location Name', 'Invoice Value Without GST', 
                    'Invoice Value With 18% GST', 'SES Number', 'SES Value', 'Invoice Tax code'];
    worksheet.addRow(headers);

    let serialNo = 1;
    let totalInvoiceWithoutGST = 0;
    let totalInvoiceWithGST = 0;

    // Process each location in processedData
    for (const [uniqueKey, data] of processedData) {
      if (data.quarter === quarter && data.totalAmount > 0) {
        // Use actualLocationCode for office lookup to get region info
        const lookupCode = data.actualLocationCode || data.locationCode;
        const office = offices.find(o => o.code === lookupCode);
        
        if (office) {
          // IMPORTANT: Use the ORIGINAL locationCode for display in report
          const displayLocationCode = data.locationCode;
          
          const invoiceValueWithoutGST = data.totalAmount;
          const invoiceValueWithGST = invoiceValueWithoutGST * 1.18;
          const taxCode = getTaxCode(lookupCode); // Use lookup code for tax determination
          
          // Use lineId from the processed data instead of incremental lineItem
          const lineItem = data.lineId || office.lineId || 10; // Fallback to office.lineId or 10

          worksheet.addRow([
            serialNo,
            lineItem,                 // Now uses actual lineId
            displayLocationCode,      // This will show 4150 for one and 41501 for the other
            data.locationName,        // This will show the correct office name
            parseFloat(invoiceValueWithoutGST.toFixed(2)),
            parseFloat(invoiceValueWithGST.toFixed(2)),
            '',                       // SES Number - left empty
            '',                       // SES Value - left empty
            taxCode
          ]);

          // Add to totals
          totalInvoiceWithoutGST += invoiceValueWithoutGST;
          totalInvoiceWithGST += invoiceValueWithGST;

          serialNo++;
        }
      }
    }

    // Add total row at the end
    worksheet.addRow([
      '',                          // S.No - empty
      '',                          // Line Item - empty
      '',                          // Location Code - empty
      'TOTAL',                     // Location Name shows "TOTAL"
      parseFloat(totalInvoiceWithoutGST.toFixed(2)),  // Total without GST
      parseFloat(totalInvoiceWithGST.toFixed(2)),     // Total with GST
      '',                          // SES Number - empty
      '',                          // SES Value - empty
      ''                           // Tax Code - empty
    ]);

    const outputDir = path.join(__dirname, '../output');
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    const fileName = `Location_Wise_Report_${quarter}.xlsx`;
    const filePath = path.join(outputDir, fileName);
    await workbook.xlsx.writeFile(filePath);

    generatedReports.locationWise = fileName;
    generatedReports.quarter = quarter;

    console.log(`✅ Location-wise report generated: ${fileName}`);
    console.log(`📊 Total Invoice Without GST: ${totalInvoiceWithoutGST.toFixed(2)}`);
    console.log(`📊 Total Invoice With GST: ${totalInvoiceWithGST.toFixed(2)}`);
    return fileName;
  } catch (error) {
    console.error('❌ Error generating location-wise report:', error);
    throw error;
  }
}


async function generateRegionWiseReport(quarter) {
  try {
    const workbook = new xlsx.Workbook();
    const worksheet = workbook.addWorksheet('Region Wise Report');

    const headers = ['Location Code', 'Location Name', 'PO NO', 'Invoice Value Without GST', 'Invoice Value With 18% GST'];
    worksheet.addRow(headers);

    // Define specific PO numbers for each location name
    const poNumbers = {
      'Southern Regional Office': '70134031',
      'Pondichery': '70134127',
      'Tamil Nadu State Office': '70157639'
    };

    // Group by region and calculate totals
    const regionTotals = new Map();

    // Process each location in processedData
    for (const [uniqueKey, data] of processedData) {
      if (data.quarter === quarter && data.totalAmount > 0) {
        const lookupCode = data.actualLocationCode || data.locationCode;
        const office = offices.find(o => o.code === lookupCode);
        
        if (office) {
          const region = office.region;
          
          if (!regionTotals.has(region)) {
            regionTotals.set(region, {
              totalAmount: 0,
              locationCode: 4000, // Use Southern Regional Office code for all regions
              locations: [] // Track individual locations in this region
            });
          }
          
          regionTotals.get(region).totalAmount += data.totalAmount;
          regionTotals.get(region).locations.push({
            locationCode: data.locationCode,
            locationName: data.locationName,
            amount: data.totalAmount
          });
        }
      }
    }

    // Add region-wise data
    let grandTotal = 0;
    let grandTotalWithGST = 0;

    for (const [region, regionData] of regionTotals) {
      const invoiceValueWithoutGST = regionData.totalAmount;
      const invoiceValueWithGST = invoiceValueWithoutGST * 1.18;
      
      // Get the specific PO number for this region, or generate a random one if not found
      let poNumber;
      if (poNumbers[region]) {
        poNumber = poNumbers[region];
      } else {
        // Fallback to random generation for regions not in the predefined list
        poNumber = generatePONumber();
      }

      worksheet.addRow([
        regionData.locationCode,
        region,
        poNumber,  // Now uses specific PO numbers
        parseFloat(invoiceValueWithoutGST.toFixed(2)),
        parseFloat(invoiceValueWithGST.toFixed(2))
      ]);

      grandTotal += invoiceValueWithoutGST;
      grandTotalWithGST += invoiceValueWithGST;
    }

    // Add grand total row
    worksheet.addRow([
      '',
      'GRAND TOTAL',
      '',
      parseFloat(grandTotal.toFixed(2)),
      parseFloat(grandTotalWithGST.toFixed(2))
    ]);

    const outputDir = path.join(__dirname, '../output');
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    const fileName = `Region_Wise_Report_${quarter}.xlsx`;
    const filePath = path.join(outputDir, fileName);
    await workbook.xlsx.writeFile(filePath);

    generatedReports.regionWise = fileName;
    generatedReports.quarter = quarter;

    console.log(`✅ Region-wise report generated: ${fileName}`);
    console.log(`📊 Grand Total Without GST: ${grandTotal.toFixed(2)}`);
    console.log(`📊 Grand Total With GST: ${grandTotalWithGST.toFixed(2)}`);
    return fileName;
  } catch (error) {
    console.error('❌ Error generating region-wise report:', error);
    throw error;
  }
}




router.post('/upload', upload.single('file'), async (req, res) => {
  try {
    const locationCode = parseInt(req.body.locationCode);
    const lineIdentifier = parseInt(req.body.lineId);
    const quarter = req.body.quarter; // Now just quarter like 'Q5'
    const plantCode = locationCode.toString();
    const [qStart, qEnd] = getQuarterRange(quarter); // Updated function call

    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }


    const workbook = new xlsx.Workbook();
    await workbook.xlsx.readFile(file.path);
    const worksheet = workbook.worksheets[0];

  const headerRow = worksheet.getRow(1);
const colMap = {};
headerRow.eachCell((cell, colNumber) => {
  const header = (cell.value || '').toString().toLowerCase().trim();
  
  // Existing mappings
  if (header.includes('short text')) colMap.shortText = colNumber;
  if (header.includes('no. of items')) colMap.noOfItems = colNumber;
  if (header.includes('service days')) colMap.serviceDays = colNumber;
  if (header.includes('amount')) colMap.amount = colNumber;
  if (header.includes('gross price')) colMap.grossPrice = colNumber;
  if (header.includes('quantity')) colMap.quantity = colNumber;
  if (header.includes('line no')) colMap.lineNo = colNumber;
  if (header.includes('remarks')) colMap.remarks = colNumber;
  
  // IMPROVED: Better consumed quantity detection
  if (header.includes('consumed quantity') || 
      header.includes('consumed qty') || 
      header.includes('consumed') ||
      header.includes('cumulative quantity') ||
      header.includes('cumulative qty')) {
    colMap.consumedQuantity = colNumber;
    console.log(`✅ Found consumed quantity column at position ${colNumber}: "${header}"`);
  }
});

// Debug: Print all detected columns
console.log('Column mapping detected:', colMap);

// ISSUE 4: Add validation after column mapping
if (!colMap.consumedQuantity) {
  console.warn('⚠️ No consumed quantity column found. Constraint will not be applied.');
  console.log('Available headers:', headerRow.values);
}
    // Add missing columns if they don't exist
    if (!colMap.amount) {
      colMap.amount = headerRow.cellCount + 1;
      headerRow.getCell(colMap.amount).value = 'Amount';
    }

    // NEW: Add remarks column if it doesn't exist
    if (!colMap.remarks) {
      colMap.remarks = headerRow.cellCount + 1;
      headerRow.getCell(colMap.remarks).value = 'Remarks';
    }

    headerRow.commit();
// First try to find exact match including lineId from form data
// You'll need to add a lineId field to your form or derive it from naming convention
let office = offices.find(o => o.code === locationCode && o.lineId === lineIdentifier);
if (!office) {
  // Fallback logic if exact match not found
  office = offices.find(o => o.code === locationCode);
}

let actualLocationCode = office ? office.code : locationCode;
let officeName = office ? sanitizeFolderName(office.name) : `Location_${locationCode}`;
let lineNo = office ? office.lineId : null;
if (office) {
  // Direct match found
  officeName = sanitizeFolderName(office.name);
  actualLocationCode = locationCode;
} else {
  // Check if this is a line-specific location code
  // Try to find if this locationCode has a corresponding office with lineId
  const lineOffice = offices.find(o => 
    o.code === locationCode && o.lineId && o.lineId !== "main"
  );
  
  if (lineOffice) {
    // This is a specific line office entry
    office = lineOffice;
    officeName = sanitizeFolderName(lineOffice.name);
    actualLocationCode = locationCode;
    // Extract line info from the office name or lineId
    if (lineOffice.lineId !== "main") {
      lineNo = lineOffice.lineId.replace('line', '');
    }
  } else {
    // Try to find base office by checking if it's a variant
    const baseCode = Math.floor(locationCode / 10) * 10; // 41501 -> 4150
    const baseOffice = offices.find(o => o.code === baseCode);
    
    if (baseOffice) {
      // This is a line variant of a base office
      office = baseOffice;
      officeName = sanitizeFolderName(baseOffice.name);
      actualLocationCode = baseCode;
      lineNo = (locationCode - baseCode).toString();
    } else {
      // No match found, use generic name
      officeName = `Location_${locationCode}`;
      actualLocationCode = locationCode;
    }
  }
}

const folderName = lineNo ? 
  `${locationCode}-${officeName}-Line${lineNo}` : 
  `${locationCode}-${officeName}`;

// Create directory paths BEFORE using them
const baseDir = path.join(__dirname, `../output/${folderName}`);
const uploadDir = path.join(baseDir, 'upload');
const calcDir = path.join(baseDir, 'calculations');

// Clean up existing directories if they exist
if (fs.existsSync(baseDir)) {
  fs.rmSync(baseDir, { recursive: true, force: true });
}

fs.mkdirSync(uploadDir, { recursive: true });
fs.mkdirSync(calcDir, { recursive: true });

const uploadWB = new xlsx.Workbook();
const uploadSheet = uploadWB.addWorksheet('Upload');

const calcWB = new xlsx.Workbook();
const calcSheet = calcWB.addWorksheet('Calculations');

uploadSheet.addRow(worksheet.getRow(1).values);
calcSheet.addRow(worksheet.getRow(1).values);

let totalAmount = 0;
let skippedRows = [];
let processedRows = 0;

// Process Excel rows here (the existing row processing logic)
  for (let i = 2; i <= worksheet.rowCount; i++) {
      const row = worksheet.getRow(i);
      const shortTextRaw = row.getCell(colMap.shortText)?.value?.toString().trim() || '';
      const quantity = parseFloat(row.getCell(colMap.quantity)?.value) || 0;
      const grossPrice = parseFloat(row.getCell(colMap.grossPrice)?.value) || 0;
      
      let uploadedConsumedQuantity = 0;
  if (colMap.consumedQuantity) {
    const consumedCell = row.getCell(colMap.consumedQuantity);
    const consumedValue = consumedCell?.value;
    
    // Debug logging
    console.log(`Row ${i}: Consumed quantity cell value:`, consumedValue, typeof consumedValue);
    
    if (consumedValue !== null && consumedValue !== undefined && consumedValue !== '') {
      uploadedConsumedQuantity = parseFloat(consumedValue) || 0;
    }
  }
  
  console.log(`Row ${i}: Uploaded consumed quantity: ${uploadedConsumedQuantity}`);

  // Skip only if all are missing
  if (!shortTextRaw && (!quantity || !grossPrice)) {
    skippedRows.push({
      rowNumber: i,
      reason: 'Missing all relevant fields (Short Text, Quantity, Gross Price)',
      shortText: shortTextRaw || 'N/A',
      quantity: quantity || 'N/A',
      grossPrice: grossPrice || 'N/A'
    });
    console.warn(`⚠️ Row ${i} skipped: Missing all relevant fields`);
    continue;
  }
      const shortText = shortTextRaw.replace(/\s+/g, ' ').replace(/\s*:\s*/g, ':').toLowerCase();
      let matchRow = null;

      // Try matching by shortText
      if (shortTextRaw) {
        try {
          const [primaryMatch] = await db.execute(
            `SELECT no_of_assets, unit_price, amc_from, amc_to FROM amc_data
             WHERE REPLACE(LOWER(REPLACE(service_short_text, ' ', '')), ':', '') = ?
               AND plant_code = ? AND quantity = ? AND total_cost = ?`,
            [shortText.replace(/\s/g, '').replace(/:/g, ''), plantCode, quantity, grossPrice]
          );
          matchRow = primaryMatch[0];
        } catch (dbError) {
          console.warn(`Database query error for row ${i}:`, dbError.message);
        }
      }

      // Fallback if no match by shortText
      if (!matchRow) {
        try {
          const [fallbackMatch] = await db.execute(
            `SELECT no_of_assets, unit_price, amc_from, amc_to FROM amc_data
             WHERE quantity = ? AND total_cost = ? AND plant_code = ?`,
            [quantity, grossPrice, plantCode]
          );
          matchRow = fallbackMatch[0];
        } catch (dbError) {
          console.warn(`Fallback database query error for row ${i}:`, dbError.message);
        }
      }

if (!matchRow) {
  skippedRows.push({
    rowNumber: i,
    reason: 'No database match found',
    shortText: shortTextRaw || 'N/A',
    quantity: quantity,
    grossPrice: grossPrice
  });
  console.warn(`⚠️ Row ${i} skipped: No DB match for Quantity=${quantity}, GrossPrice=${grossPrice}`);
  continue;
}

processedRows++;

const noOfItems = matchRow.no_of_assets || 0;
const unitPrice = matchRow.unit_price || 0;
const serviceDays = calculateServiceDays(matchRow.amc_from, matchRow.amc_to, qStart, qEnd);

// Get uploaded consumed quantity from Excel

if (colMap.consumedQuantity) {
  const consumedCell = row.getCell(colMap.consumedQuantity);
  const consumedValue = consumedCell?.value;
  
  if (consumedValue !== null && consumedValue !== undefined && consumedValue !== '') {
    uploadedConsumedQuantity = parseFloat(consumedValue) || 0;
  }
}

// Calculate expected consumed quantity from previous quarters
const previousQuartersData = calculateConsumedQuantityFromPreviousQuarters(
  matchRow.amc_from, 
  matchRow.amc_to, 
  quarter, 
  noOfItems
);

const calculatedConsumedQuantity = previousQuartersData.totalConsumedQuantity;
const currentQuarterConsumption = noOfItems * serviceDays;

// Expected total consumed quantity (previous quarters + current quarter)
const expectedTotalConsumedQuantity = calculatedConsumedQuantity;

console.log(`Row ${i}: Previous quarters consumed: ${calculatedConsumedQuantity}`);
console.log(`Row ${i}: Current quarter consumption: ${currentQuarterConsumption}`);
console.log(`Row ${i}: Expected total consumed: ${expectedTotalConsumedQuantity}`);
console.log(`Row ${i}: Uploaded consumed quantity: ${uploadedConsumedQuantity}`);

let amount = noOfItems * serviceDays * unitPrice;
let remarks = '';
let adjustedAmount = amount;

// Enhanced consumed quantity validation and adjustment
if (uploadedConsumedQuantity > 0 && serviceDays > 0) {
  // Check if there's a discrepancy between uploaded and expected consumed quantity
  const discrepancy = uploadedConsumedQuantity - expectedTotalConsumedQuantity;
  
  if (Math.abs(discrepancy) > 0.01) { // Using small tolerance for floating point comparison
    console.log(`Row ${i}: Discrepancy detected: ${discrepancy}`);
    
    if (discrepancy > 0) {
      // Over-consumption: The uploaded consumed quantity is higher than expected
      // This means less should be charged in the current quarter
      
      const excessQuantity = discrepancy;
      const adjustedCurrentQuarterConsumption = Math.max(0, currentQuarterConsumption - excessQuantity);
      adjustedAmount = adjustedCurrentQuarterConsumption * unitPrice;
      
      remarks = `Consumed quantity adjustment: Uploaded=${uploadedConsumedQuantity}, Expected=${expectedTotalConsumedQuantity}, Excess=${excessQuantity.toFixed(0)}. Current quarter consumption adjusted from ${currentQuarterConsumption} to ${adjustedCurrentQuarterConsumption}.`;
      
      console.log(`🔄 Row ${i}: Over-consumption detected. Adjusted consumption: ${adjustedCurrentQuarterConsumption}`);
      
    } else {
      // Under-consumption: The uploaded consumed quantity is lower than expected
      // This means more should be charged in the current quarter to catch up
      
      const shortfallQuantity = Math.abs(discrepancy);
      const adjustedCurrentQuarterConsumption = currentQuarterConsumption + shortfallQuantity;
      adjustedAmount = adjustedCurrentQuarterConsumption * unitPrice;
      
      remarks = `Consumed quantity adjustment: Uploaded=${uploadedConsumedQuantity}, Expected=${expectedTotalConsumedQuantity}, Shortfall=${shortfallQuantity.toFixed(0)}. Current quarter consumption adjusted from ${currentQuarterConsumption} to ${adjustedCurrentQuarterConsumption}.`;
      
      console.log(`🔄 Row ${i}: Under-consumption detected. Adjusted consumption: ${adjustedCurrentQuarterConsumption}`);
    }
    
    // Ensure adjusted amount is not negative
    if (adjustedAmount < 0) {
      adjustedAmount = 0;
      remarks += ' (Amount capped at 0 due to negative adjustment)';
    }
    
    console.log(`💰 Row ${i}: Amount adjustment - Original: ${amount.toFixed(2)}, Adjusted: ${adjustedAmount.toFixed(2)}`);
    
  } else {
    // No significant discrepancy
    console.log(`✅ Row ${i}: Consumed quantity matches expected (diff: ${discrepancy.toFixed(2)})`);
  }
} else {
  // No consumed quantity provided or no service days
  if (uploadedConsumedQuantity === 0) {
    console.log(`ℹ️ Row ${i}: No consumed quantity provided in upload`);
  }
  if (serviceDays === 0) {
    console.log(`ℹ️ Row ${i}: No service days for current quarter (outside AMC range)`);
  }
}

// Additional validation: Check if AMC period is still active
const amcFromDate = new Date(matchRow.amc_from);
const amcToDate = new Date(matchRow.amc_to);
const currentQuarterStart = new Date(qStart);
const currentQuarterEnd = new Date(qEnd);

if (amcToDate < currentQuarterStart) {
  // AMC period has ended before current quarter
  adjustedAmount = 0;
  if (remarks) {
    remarks += ' AMC period ended before current quarter.';
  } else {
    remarks = 'AMC period ended before current quarter.';
  }
  console.log(`⚠️ Row ${i}: AMC period ended before current quarter. Amount set to 0.`);
}

// Use adjusted amount for total calculation
totalAmount += adjustedAmount;

// Update row values
if (colMap.noOfItems) row.getCell(colMap.noOfItems).value = noOfItems;
if (colMap.serviceDays) row.getCell(colMap.serviceDays).value = serviceDays;
row.getCell(colMap.amount).value = parseFloat(adjustedAmount.toFixed(2));

// Add remarks only if there's an adjustment or issue
if (colMap.remarks && remarks) {
  row.getCell(colMap.remarks).value = remarks;
}

row.commit();

  // For upload sheet (without amount column)
  const uploadRow = row.values.map((val, idx) =>
    idx === colMap.amount ? null : val
  );
  uploadSheet.addRow(uploadRow);
  calcSheet.addRow(row.values);
}


// Add total row to both worksheets
if (totalAmount > 0) {
  const uploadTotalRow = new Array(headerRow.cellCount).fill(null);
  uploadTotalRow[0] = 'TOTAL';
  uploadSheet.addRow(uploadTotalRow);

  const calcTotalRow = new Array(headerRow.cellCount).fill(null);
  calcTotalRow[0] = 'TOTAL';
  calcTotalRow[colMap.amount - 1] = totalAmount;
  calcSheet.addRow(calcTotalRow);
}

const uploadPath = path.join(uploadDir, `upload_${folderName}.xlsx`);
const calcPath = path.join(calcDir, `calculated_${folderName}.xlsx`);

await uploadWB.xlsx.writeFile(uploadPath);
await calcWB.xlsx.writeFile(calcPath);

const zipFile = `${folderName}.zip`;
const zipPath = path.join(__dirname, '../output', zipFile);

// Remove existing zip if it exists
if (fs.existsSync(zipPath)) {
  fs.unlinkSync(zipPath);
}

const output = fs.createWriteStream(zipPath);
const archive = archiver('zip', { zlib: { level: 9 } });

archive.pipe(output);
archive.directory(baseDir, false);
await archive.finalize();

const uniqueKey = `${locationCode}_${lineIdentifier}_${quarter}`;

    processedData.set(uniqueKey, {
      locationCode,
      actualLocationCode,
      lineId: lineIdentifier,
      lineNumber: lineNo,
      locationName: office ? office.name : officeName,
      quarter, // Now just stores quarter like 'Q5'
      totalAmount,
      folderName,
      processedAt: new Date(),
      folderPath: baseDir,
      totalRows: worksheet.rowCount - 1,
      processedRows,
      skippedRows
    });

    // Clean up
    fs.unlinkSync(file.path);
    
    console.log(`✅ File processed successfully for location ${locationCode}, Quarter ${quarter}, Total Amount: ${totalAmount}`);
    res.json({ success: true, fileName: zipFile, totalAmount: totalAmount });
  } catch (err) {
    console.error('❌ Error in /upload route:', err);
    
    // Clean up file if it exists
    if (req.file && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }
    
    res.status(500).json({ error: 'Internal Server Error', details: err.message });
  }
});
// Generate Location-wise Report endpoint
router.post('/generate-location-report', async (req, res) => {
  try {
    const { quarter } = req.body; // Only quarter needed now
    
    if (!quarter) {
      return res.status(400).json({ error: 'Quarter is required' });
    }

    // Check if there's any processed data for the given quarter
    const hasData = Array.from(processedData.values()).some(
      data => data.quarter === quarter && data.totalAmount > 0
    );

    if (!hasData) {
      return res.status(400).json({ 
        error: `No processed data found for quarter ${quarter}. Please upload location files first.` 
      });
    }

    const fileName = await generateLocationWiseReport(quarter);
    res.json({ success: true, fileName: fileName });
  } catch (err) {
    console.error('❌ Error in /upload route:', err);
    
    // Clean up file if it exists
    if (req.file && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }
    
    res.status(500).json({ error: 'Internal Server Error', details: err.message });
  }
});

// Generate Region-wise Report endpoint
router.post('/generate-region-report', async (req, res) => {
  try {
    const { quarter } = req.body; // Only quarter needed now
    
    if (!quarter) {
      return res.status(400).json({ error: 'Quarter is required' });
    }

    // Check if there's any processed data for the given quarter
    const hasData = Array.from(processedData.values()).some(
      data => data.quarter === quarter && data.totalAmount > 0
    );

    if (!hasData) {
      return res.status(400).json({ 
        error: `No processed data found for quarter ${quarter}. Please upload location files first.` 
      });
    }

    const fileName = await generateRegionWiseReport(quarter);
    res.json({ success: true, fileName: fileName });
  } catch (error) {
    console.error('❌ Error generating region-wise report:', error);
    res.status(500).json({ error: 'Failed to generate region-wise report', details: error.message });
  }
});

// FIXED download-all endpoint
router.get('/download-all', async (req, res) => {
  try {
    const outputDir = path.join(__dirname, '../output');
    
    if (!fs.existsSync(outputDir)) {
      return res.status(400).json({ error: 'Output directory not found' });
    }

    // Check if both reports have been generated
    if (!generatedReports.locationWise || !generatedReports.regionWise) {
      return res.status(400).json({ 
        error: 'Please generate both location-wise and region-wise reports first.',
        availableReports: {
          locationWise: generatedReports.locationWise,
          regionWise: generatedReports.regionWise
        }
      });
    }

    // Verify the report files exist
    const locationReportPath = path.join(outputDir, generatedReports.locationWise);
    const regionReportPath = path.join(outputDir, generatedReports.regionWise);

    if (!fs.existsSync(locationReportPath) || !fs.existsSync(regionReportPath)) {
      return res.status(400).json({ 
        error: 'Required report files not found. Please regenerate the reports.',
        missingFiles: {
          locationReport: !fs.existsSync(locationReportPath),
          regionReport: !fs.existsSync(regionReportPath)
        }
      });
    }

    const quarter = generatedReports.quarter; // Now just quarter like 'Q5'
    
    // Create main folder name with timestamp
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const mainFolderName = `Invoices_and_Reports_${quarter}_${timestamp}`;
    const mainFolderPath = path.join(outputDir, mainFolderName);
    const invoicesSubfolder = path.join(mainFolderPath, 'Location_Invoices');
    const reportsSubfolder = path.join(mainFolderPath, 'Reports');

    // Clean up if exists
    if (fs.existsSync(mainFolderPath)) {
      fs.rmSync(mainFolderPath, { recursive: true, force: true });
    }

    // Create directory structure
    fs.mkdirSync(invoicesSubfolder, { recursive: true });
    fs.mkdirSync(reportsSubfolder, { recursive: true });

    // Copy only the location folders for processed data in this quarter
    let copiedFolders = 0;
    const processedLocations = [];

    for (const [uniqueKey, data] of processedData) {
      if (data.quarter === quarter && data.totalAmount > 0) {
        const sourceFolderPath = path.join(outputDir, data.folderName);
        
        if (fs.existsSync(sourceFolderPath)) {
          const destPath = path.join(invoicesSubfolder, data.folderName);
          
          try {
            fs.cpSync(sourceFolderPath, destPath, { recursive: true });
            copiedFolders++;
            processedLocations.push({
              locationCode: data.locationCode,
              actualLocationCode: data.actualLocationCode,
              lineNumber: data.lineNumber,
              lineId: data.lineId,
              locationName: data.locationName,
              totalAmount: data.totalAmount,
              folderName: data.folderName
            });
            console.log(`✅ Copied folder: ${data.folderName}`);
          } catch (copyError) {
            console.warn(`⚠️ Failed to copy folder ${data.folderName}:`, copyError.message);
          }
        } else {
          console.warn(`⚠️ Source folder not found: ${sourceFolderPath}`);
        }
      }
    }

    if (copiedFolders === 0) {
      return res.status(400).json({ 
        error: `No location folders found for quarter ${quarter}. Please upload location files first.` 
      });
    }
    // Copy the EXACT reports that were generated
     try {
      fs.copyFileSync(locationReportPath, path.join(reportsSubfolder, generatedReports.locationWise));
      fs.copyFileSync(regionReportPath, path.join(reportsSubfolder, generatedReports.regionWise));
      console.log(`✅ Copied reports to ${reportsSubfolder}`);
    } catch (reportError) {
      console.error('❌ Error copying reports:', reportError);
      return res.status(500).json({ error: 'Failed to copy reports' });
    }

    // Updated summaryData object
    const summaryData = {
      generated_at: new Date().toISOString(),
      quarter: quarter, // Now just quarter like 'Q5'
      total_location_folders: copiedFolders,
      reports_included: [generatedReports.locationWise, generatedReports.regionWise],
      structure: {
        main_folder: mainFolderName,
        invoices_folder: 'Location_Invoices',
        reports_folder: 'Reports'
      },
      processing_summary: {
        total_locations_processed: processedLocations.length,
        total_rows_across_all_files: 0,
        total_successful_rows: 0,
        total_skipped_rows: 0,
        total_amount_calculated: 0
      },
      processed_locations: processedLocations.map(loc => {
        // Find the correct data entry by matching location details
        let matchingData = null;
        for (const [key, data] of processedData) {
          if (data.locationCode === loc.locationCode && 
              data.lineNumber === loc.lineNumber &&
              data.quarter === quarter) {
            matchingData = data;
            break;
          }
        }
        
        return {
          location_code: loc.locationCode,
          actual_location_code: loc.actualLocationCode,
          line_id: loc.lineId,
          line_number: loc.lineNumber,
          location_name: loc.locationName,
          folder_name: loc.folderName,
          total_amount_calculated: loc.totalAmount,
          processing_stats: {
            total_rows_in_file: matchingData?.totalRows || 0,
            successfully_processed_rows: matchingData?.processedRows || 0,
            skipped_rows_count: matchingData?.skippedRows?.length || 0,
            processing_success_rate: matchingData?.totalRows > 0 ? 
              `${((matchingData?.processedRows || 0) / matchingData?.totalRows * 100).toFixed(1)}%` : '0%'
          },
          errors_and_skipped_rows: (matchingData?.skippedRows || []).map(skip => ({
            row_number: skip.rowNumber,
            error_reason: skip.reason,
            row_data: {
              short_text: skip.shortText,
              quantity: skip.quantity,
              gross_price: skip.grossPrice
            }
          }))
        };
      }),
      error_summary: {
        locations_with_errors: 0,
        common_error_types: {},
        total_errors_across_all_files: 0
      }
    };
// Calculate summary statistics and error analysis
let totalRowsAll = 0;
let totalSuccessfulAll = 0;
let totalSkippedAll = 0;
let totalAmountAll = 0;
let locationsWithErrors = 0;
const errorTypes = {};

summaryData.processed_locations.forEach(loc => {
  totalRowsAll += loc.processing_stats.total_rows_in_file;
  totalSuccessfulAll += loc.processing_stats.successfully_processed_rows;
  totalSkippedAll += loc.processing_stats.skipped_rows_count;
  totalAmountAll += loc.total_amount_calculated;
  
  if (loc.errors_and_skipped_rows.length > 0) {
    locationsWithErrors++;
    
    // Count error types
    loc.errors_and_skipped_rows.forEach(error => {
      const errorType = error.error_reason;
      if (errorTypes[errorType]) {
        errorTypes[errorType]++;
      } else {
        errorTypes[errorType] = 1;
      }
    });
  }
});

// Update the summary statistics
summaryData.processing_summary.total_rows_across_all_files = totalRowsAll;
summaryData.processing_summary.total_successful_rows = totalSuccessfulAll;
summaryData.processing_summary.total_skipped_rows = totalSkippedAll;
summaryData.processing_summary.total_amount_calculated = parseFloat(totalAmountAll.toFixed(2));
summaryData.processing_summary.overall_success_rate = totalRowsAll > 0 ? 
  `${(totalSuccessfulAll / totalRowsAll * 100).toFixed(1)}%` : '0%';

// Update error summary
summaryData.error_summary.locations_with_errors = locationsWithErrors;
summaryData.error_summary.total_errors_across_all_files = totalSkippedAll;
summaryData.error_summary.common_error_types = errorTypes;

// Add error recommendations
summaryData.error_summary.recommendations = [];
if (errorTypes['No database match found']) {
  summaryData.error_summary.recommendations.push(
    "Some rows couldn't be matched with database records. Verify that the plant code, quantity, and gross price values in the Excel file match the database entries."
  );
}
if (errorTypes['Missing all relevant fields (Short Text, Quantity, Gross Price)']) {
  summaryData.error_summary.recommendations.push(
    "Some rows have missing critical data. Ensure all rows have Short Text, Quantity, and Gross Price filled in."
  );
}

// Write the enhanced summary
 fs.writeFileSync(
      path.join(mainFolderPath, 'Summary.json'), 
      JSON.stringify(summaryData, null, 2)
    );

    // Create the final zip
    const finalZipName = `${mainFolderName}.zip`;
    const finalZipPath = path.join(outputDir, finalZipName);
    
    if (fs.existsSync(finalZipPath)) {
      fs.unlinkSync(finalZipPath);
    }

    const output = fs.createWriteStream(finalZipPath);
    const archive = archiver('zip', { zlib: { level: 9 } });

    output.on('close', () => {
      console.log(`✅ Created final zip: ${finalZipName} (${archive.pointer()} bytes)`);
      
      // Clean up temporary folder
      fs.rmSync(mainFolderPath, { recursive: true, force: true });
      
      res.download(finalZipPath, finalZipName, (err) => {
        if (err) {
          console.error('❌ Error sending file:', err);
        } else {
          console.log(`✅ File download initiated: ${finalZipName}`);
          // Clean up the zip file after a delay
          setTimeout(() => {
            if (fs.existsSync(finalZipPath)) {
              fs.unlinkSync(finalZipPath);
              console.log(`🗑️ Cleaned up zip file: ${finalZipName}`);
            }
          }, 5000);
        }
      });
    });

    output.on('error', (err) => {
      console.error('❌ Error creating zip:', err);
      res.status(500).json({ error: 'Failed to create zip file' });
    });

    archive.on('error', (err) => {
      console.error('❌ Archive error:', err);
      res.status(500).json({ error: 'Failed to archive files' });
    });

    archive.pipe(output);
    archive.directory(mainFolderPath, false);
    await archive.finalize();

  } catch (err) {
    console.error('❌ Error in /download-all:', err);
    res.status(500).json({ error: 'Failed to generate structured download', details: err.message });
  }
});

// Updated processed data summary endpoint
router.get('/processed-data', (req, res) => {
  const summary = Array.from(processedData.entries()).map(([uniqueKey, data]) => ({
    locationCode: data.locationCode,
    lineId: data.lineId,
    locationName: data.locationName,
    quarter: data.quarter, // Now just quarter like 'Q5'
    totalAmount: data.totalAmount,
    processedAt: data.processedAt
  }));

  res.json({ success: true, data: summary });
});
// Get processed data summary
router.get('/processed-data', (req, res) => {
  const summary = Array.from(processedData.entries()).map(([uniqueKey, data]) => ({
    locationCode: data.locationCode,
    lineId: data.lineId,
    locationName: data.locationName,
    quarter: data.quarter,
    yearRange: data.yearRange,
    totalAmount: data.totalAmount,
    processedAt: data.processedAt
  }));

  res.json({ success: true, data: summary });
});


module.exports = router;