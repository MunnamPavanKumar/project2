const express = require('express');
const router = express.Router();
const multer = require('multer');
const xlsx = require('exceljs');
const db = require('../db/connection');
const fs = require('fs');
const path = require('path');
const archiver = require('archiver');

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
  
  return ranges[quarter] || ranges['Q1']; // Default to Q1 if quarter not found
}

function calculateServiceDays(amcFrom, amcTo, qStart, qEnd) {
  const from = new Date(amcFrom);
  const to = new Date(amcTo);
  const start = new Date(qStart);
  const end = new Date(qEnd);
  
  // Check if ranges don't overlap
  if (to < start || from > end) return 0;
  
  // Find the actual overlapping period
  const actualStart = from > start ? from : start;
  const actualEnd = to < end ? to : end;
  
  // Normalize dates to remove time components for accurate day calculation
  const normalizedStart = new Date(actualStart.getFullYear(), actualStart.getMonth(), actualStart.getDate());
  const normalizedEnd = new Date(actualEnd.getFullYear(), actualEnd.getMonth(), actualEnd.getDate());
  
  // Calculate days difference (including both start and end dates)
  const timeDiff = normalizedEnd.getTime() - normalizedStart.getTime();
  const daysDiff = Math.floor(timeDiff / (1000 * 60 * 60 * 24)) + 1;
  
  return daysDiff;
}
let consumedQuantityTracker = new Map();

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


function getQuarterNumber(quarter) {
  return parseInt(quarter.replace('Q', ''));
}


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

// Updated store generated reports object
let generatedReports = {
  locationWise: null,
  regionWise: null,
  quarter: null
};

// New function to generate consumed quantity validation report
async function generateConsumedQuantityReport(quarter) {
  try {
    const workbook = new xlsx.Workbook();
    const worksheet = workbook.addWorksheet('Consumed Quantity Validation');

    const headers = [
      'Location Code', 'Short Text', 'Quarter', 'Consumed Qty (File)', 
      'Expected Previous Quarters Qty', 'Current Quarter Expected', 'Discrepancy', 
      'Service Days Adjustment', 'Amount Adjustment', 'Base Amount', 'Final Amount', 'Status'
    ];
    worksheet.addRow(headers);

    let totalDiscrepancies = 0;
    let totalAdjustments = 0;
    let totalBaseAmount = 0;
    let totalFinalAmount = 0;

    for (const [trackingKey, tracker] of consumedQuantityTracker) {
      if (tracker.quarterlyData.has(quarter)) {
        const quarterData = tracker.quarterlyData.get(quarter);
        const hasDiscrepancy = Math.abs(quarterData.discrepancy) > 0.01;
        
        // Find the corresponding processed data to get amounts
        let baseAmount = 0;
        let finalAmount = 0;
        
        for (const [uniqueKey, processedItem] of processedData) {
          if (processedItem.locationCode === tracker.locationCode && 
              processedItem.quarter === quarter) {
            baseAmount = processedItem.baseAmount || 0;
            finalAmount = processedItem.totalAmount || 0;
            break;
          }
        }
        
        const adjustmentAmount = finalAmount - baseAmount;
        const serviceDaysAdjustment = quarterData.discrepancy / quarterData.noOfItems;
        
        if (hasDiscrepancy) {
          totalDiscrepancies++;
          totalAdjustments += adjustmentAmount;
        }

        totalBaseAmount += baseAmount;
        totalFinalAmount += finalAmount;

        worksheet.addRow([
          tracker.locationCode,
          tracker.shortText,
          quarter,
          quarterData.consumedQtyFromFile,
          quarterData.expectedCumulativeConsumedQty.toFixed(2),
          quarterData.currentQuarterExpected.toFixed(2),
          quarterData.discrepancy.toFixed(2),
          serviceDaysAdjustment.toFixed(2),
          adjustmentAmount.toFixed(2),
          baseAmount.toFixed(2),
          finalAmount.toFixed(2),
          hasDiscrepancy ? 'ADJUSTED' : 'OK'
        ]);
      }
    }

    // Add summary row
    worksheet.addRow([
      '', '', 'SUMMARY', '', '', '', '',
      '', // Service days adjustment
      totalAdjustments.toFixed(2),
      totalBaseAmount.toFixed(2),
      totalFinalAmount.toFixed(2),
      `${totalDiscrepancies} discrepancies found`
    ]);

    const outputDir = path.join(__dirname, '../output');
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    const fileName = `Consumed_Quantity_Validation_${quarter}.xlsx`;
    const filePath = path.join(outputDir, fileName);
    await workbook.xlsx.writeFile(filePath);

    console.log(`✅ Consumed quantity validation report generated: ${fileName}`);
    console.log(`📊 Total Discrepancies: ${totalDiscrepancies}`);
    console.log(`📊 Total Adjustments: ${totalAdjustments.toFixed(2)}`);
    return fileName;
  } catch (error) {
    console.error('❌ Error generating consumed quantity report:', error);
    throw error;
  }
}
 function validateConsumedQuantityEnhanced(locationCode, shortText, currentQuarter, consumedQtyFromFile, noOfItems, serviceDays, unitPrice, amcFrom, amcTo) {
  const trackingKey = `${locationCode}_${shortText}`;
  
  if (!consumedQuantityTracker.has(trackingKey)) {
    consumedQuantityTracker.set(trackingKey, {
      locationCode,
      shortText,
      quarterlyData: new Map(),
      lastProcessedQuarter: null
    });
  }

  const tracker = consumedQuantityTracker.get(trackingKey);
  const currentQuarterNum = getQuarterNumber(currentQuarter);
  
  // Calculate expected cumulative consumed quantity for PREVIOUS quarters only
  let expectedCumulativeConsumedQty = 0;
  
  

// Calculate expected consumption for each quarter BEFORE current quarter
for (let q = 1; q < currentQuarterNum; q++) {
  const quarterKey = `Q${q}`;
  const [qStart, qEnd] = getQuarterRange(quarterKey);
  const quarterServiceDays = calculateServiceDays(amcFrom, amcTo, qStart, qEnd);
  expectedCumulativeConsumedQty += noOfItems * quarterServiceDays;
}

  
  // The consumed quantity from file should equal expected cumulative from previous quarters
 // Compare uploaded cumulative Q1–Q4 quantity with expected Q1–Q4 consumption
const discrepancy = consumedQtyFromFile - expectedCumulativeConsumedQty;

// Q5 expected quantity (for adjustment only)
const currentQuarterExpected = noOfItems * serviceDays;

// No need to subtract anything from consumedQtyFromFile to calculate Q5 consumption
// because Q5 is NOT included in the uploaded value

  
  // Store current quarter data
  tracker.quarterlyData.set(currentQuarter, {
    consumedQtyFromFile,
    expectedCumulativeConsumedQty,
    currentQuarterExpected,
    discrepancy: discrepancy,
    noOfItems,
    serviceDays,
    unitPrice
  });
  
  tracker.lastProcessedQuarter = currentQuarter;
  
  // Calculate adjustment for current quarter
  let adjustmentAmount = 0;
let adjustedServiceDays = serviceDays;

if (Math.abs(discrepancy) > 0.01) {
  // Directly calculate amount difference instead of adjusting days
  adjustmentAmount = discrepancy * unitPrice;
}

  
  return {
    isValid: Math.abs(discrepancy) <= 0.01,
    discrepancy: discrepancy,
    consumedQtyFromFile: consumedQtyFromFile,
    expectedConsumedQty: expectedCumulativeConsumedQty,
    currentQuarterExpected: currentQuarterExpected,
    adjustmentAmount: adjustmentAmount,
    adjustedServiceDays: adjustedServiceDays,
    originalServiceDays: serviceDays
  };
}module.exports = router;
