router.post('/upload', upload.single('file'), async (req, res) => {
  try {
    const { locationCode, lineId, quarter } = validateRequestData(req);
    const file = req.file;
    
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    const plantCode = locationCode.toString();
    const [qStart, qEnd] = getQuarterRange(quarter);

    // Process Excel file
    const workbook = new xlsx.Workbook();
    await workbook.xlsx.readFile(file.path);
    const worksheet = workbook.worksheets[0];

    // Map column headers
    const colMap = mapColumnHeaders(worksheet);
    
    // Find office information
    const officeInfo = findOfficeInfo(locationCode, lineId);
    
    // Create output directories
    const paths = createOutputDirectories(officeInfo);
    
    // Create workbooks for output
    const { uploadWB, calcWB } = createOutputWorkbooks(worksheet, colMap);
    
    // Process rows and calculate amounts
    const processingResults = await processWorksheetRows(
      worksheet, 
      colMap, 
      plantCode, 
      qStart, 
      qEnd, 
      locationCode, 
      quarter,
      uploadWB.worksheets[0], 
      calcWB.worksheets[0]
    );

    // Add total rows
    addTotalRows(uploadWB.worksheets[0], calcWB.worksheets[0], processingResults.totalAmount, colMap);

    // Save files and create zip
    const zipFile = await saveFilesAndCreateZip(paths, uploadWB, calcWB, officeInfo.folderName);

    // Store processed data
    storeProcessedData(locationCode, lineId, quarter, officeInfo, processingResults);

    // Clean up temporary files
    cleanupTempFiles(file.path, paths.baseDir);
    
    console.log(`✅ File processed successfully for location ${locationCode}, Quarter ${quarter}, Total Amount: ${processingResults.totalAmount}`);
    res.json({ 
      success: true, 
      fileName: zipFile, 
      totalAmount: processingResults.totalAmount,
      processedRows: processingResults.processedRows,
      skippedRows: processingResults.skippedRows.length
    });

  } catch (err) {
    console.error('❌ Error in /upload route:', err);
    
    // Clean up file if it exists
    if (req.file && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }
    
    res.status(500).json({ 
      error: 'Internal Server Error', 
      details: process.env.NODE_ENV === 'development' ? err.message : 'Processing failed'
    });
  }
});

// Helper functions
function validateRequestData(req) {
  const locationCode = parseInt(req.body.locationCode);
  const lineId = parseInt(req.body.lineId);
  const quarter = req.body.quarter;

  if (!locationCode || !lineId || !quarter) {
    throw new Error('Missing required parameters: locationCode, lineId, or quarter');
  }

  return { locationCode, lineId, quarter };
}

function mapColumnHeaders(worksheet) {
  const headerRow = worksheet.getRow(1);
  const colMap = {};
  
  headerRow.eachCell((cell, colNumber) => {
    const header = (cell.value || '').toString().toLowerCase().trim();
    if (header.includes('short text')) colMap.shortText = colNumber;
    if (header.includes('no. of items')) colMap.noOfItems = colNumber;
    if (header.includes('service days')) colMap.serviceDays = colNumber;
    if (header.includes('amount')) colMap.amount = colNumber;
    if (header.includes('gross price')) colMap.grossPrice = colNumber;
    if (header.includes('quantity')) colMap.quantity = colNumber;
    if (header.includes('line no')) colMap.lineNo = colNumber;
    if (header.includes('consumed qty')) colMap.consumedQty = colNumber;
    if (header.includes('remarks')) colMap.remarks = colNumber;
  });

  // Add amount column if not present
  if (!colMap.amount) {
    colMap.amount = headerRow.cellCount + 1;
    headerRow.getCell(colMap.amount).value = 'Amount';
    headerRow.commit();
  }

  return colMap;
}

function findOfficeInfo(locationCode, lineIdentifier) {
  // First try to find exact match
  let office = offices.find(o => o.code === locationCode && o.lineId === lineIdentifier);
  
  if (!office) {
    office = offices.find(o => o.code === locationCode);
  }

  let actualLocationCode = office ? office.code : locationCode;
  let officeName = office ? sanitizeFolderName(office.name) : `Location_${locationCode}`;
  let lineNo = office ? office.lineId : null;

  if (!office) {
    // Check for line-specific location code
    const lineOffice = offices.find(o => 
      o.code === locationCode && o.lineId && o.lineId !== "main"
    );
    
    if (lineOffice) {
      office = lineOffice;
      officeName = sanitizeFolderName(lineOffice.name);
      actualLocationCode = locationCode;
      lineNo = lineOffice.lineId !== "main" ? lineOffice.lineId.replace('line', '') : null;
    } else {
      // Try base code variant
      const baseCode = Math.floor(locationCode / 10) * 10;
      const baseOffice = offices.find(o => o.code === baseCode);
      
      if (baseOffice) {
        office = baseOffice;
        officeName = sanitizeFolderName(baseOffice.name);
        actualLocationCode = baseCode;
        lineNo = (locationCode - baseCode).toString();
      }
    }
  }

  const folderName = lineNo ? 
    `${locationCode}-${officeName}-Line${lineNo}` : 
    `${locationCode}-${officeName}`;

  return {
    office,
    actualLocationCode,
    officeName,
    lineNo,
    folderName
  };
}

function createOutputDirectories(officeInfo) {
  const baseDir = path.join(__dirname, `../output/${officeInfo.folderName}`);
  const uploadDir = path.join(baseDir, 'upload');
  const calcDir = path.join(baseDir, 'calculations');

  // Clean up existing directories
  if (fs.existsSync(baseDir)) {
    fs.rmSync(baseDir, { recursive: true, force: true });
  }

  fs.mkdirSync(uploadDir, { recursive: true });
  fs.mkdirSync(calcDir, { recursive: true });

  return { baseDir, uploadDir, calcDir };
}

function createOutputWorkbooks(worksheet, colMap) {
  const uploadWB = new xlsx.Workbook();
  const uploadSheet = uploadWB.addWorksheet('Upload');

  const calcWB = new xlsx.Workbook();
  const calcSheet = calcWB.addWorksheet('Calculations');

  // Copy headers
  uploadSheet.addRow(worksheet.getRow(1).values);
  calcSheet.addRow(worksheet.getRow(1).values);

  return { uploadWB, calcWB };
}

async function processWorksheetRows(worksheet, colMap, plantCode, qStart, qEnd, locationCode, quarter, uploadSheet, calcSheet) {
  let totalAmount = 0;
  let totalBaseAmount = 0;
  let totalAdjustmentAmount = 0;
  let skippedRows = [];
  let processedRows = 0;
  let consumedQtyValidations = [];

  for (let i = 2; i <= worksheet.rowCount; i++) {
    const row = worksheet.getRow(i);
    const rowData = extractRowData(row, colMap);

    // Skip if missing essential data
    if (!rowData.shortTextRaw && (!rowData.quantity || !rowData.grossPrice)) {
      skippedRows.push({
        rowNumber: i,
        reason: 'Missing all relevant fields',
        shortText: rowData.shortTextRaw || 'N/A',
        quantity: rowData.quantity || 'N/A',
        grossPrice: rowData.grossPrice || 'N/A'
      });
      continue;
    }

    // Find matching database record
    const matchRow = await findDatabaseMatch(rowData, plantCode, i);
    
    if (!matchRow) {
      skippedRows.push({
        rowNumber: i,
        reason: 'No database match found',
        shortText: rowData.shortTextRaw || 'N/A',
        quantity: rowData.quantity,
        grossPrice: rowData.grossPrice
      });
      continue;
    }

    // Calculate amounts and validate
    const calculationResult = calculateRowAmount(
      matchRow, 
      rowData, 
      qStart, 
      qEnd, 
      locationCode, 
      quarter
    );

    // Track consumed quantity validations
    if (calculationResult.validation) {
      consumedQtyValidations.push({
        rowNumber: i,
        shortText: rowData.shortTextRaw,
        ...calculationResult.validation
      });
    }

    // Update row with calculated values
    updateRowWithCalculations(row, colMap, calculationResult);
    
    // Add to output sheets
    const uploadRow = row.values.map((val, idx) =>
      idx === colMap.amount ? null : val
    );
    uploadSheet.addRow(uploadRow);
    calcSheet.addRow(row.values);

    totalAmount += calculationResult.amount;
    totalBaseAmount += calculationResult.baseAmount;
    totalAdjustmentAmount += calculationResult.adjustmentAmount;
    processedRows++;
  }

  return { 
    totalAmount, 
    baseAmount: totalBaseAmount,
    adjustmentAmount: totalAdjustmentAmount,
    skippedRows, 
    processedRows,
    consumedQtyValidations
  };
}
function extractRowData(row, colMap) {
  return {
    shortTextRaw: row.getCell(colMap.shortText)?.value?.toString().trim() || '',
    quantity: parseFloat(row.getCell(colMap.quantity)?.value) || 0,
    grossPrice: parseFloat(row.getCell(colMap.grossPrice)?.value) || 0,
    consumedQty: parseFloat(row.getCell(colMap.consumedQty)?.value) || 0
  };
}

async function findDatabaseMatch(rowData, plantCode, rowNumber) {
  const shortText = rowData.shortTextRaw
    .replace(/\s+/g, ' ')
    .replace(/\s*:\s*/g, ':')
    .toLowerCase();

  try {
    // Try primary match by shortText
    if (rowData.shortTextRaw) {
      const [primaryMatch] = await db.execute(
        `SELECT no_of_assets, unit_price, amc_from, amc_to FROM amc_data
         WHERE REPLACE(LOWER(REPLACE(service_short_text, ' ', '')), ':', '') = ?
           AND plant_code = ? AND quantity = ? AND total_cost = ?`,
        [shortText.replace(/\s/g, '').replace(/:/g, ''), plantCode, rowData.quantity, rowData.grossPrice]
      );
      
      if (primaryMatch[0]) return primaryMatch[0];
    }

    // Fallback match
    const [fallbackMatch] = await db.execute(
      `SELECT no_of_assets, unit_price, amc_from, amc_to FROM amc_data
       WHERE quantity = ? AND total_cost = ? AND plant_code = ?`,
      [rowData.quantity, rowData.grossPrice, plantCode]
    );
    
    return fallbackMatch[0] || null;
  } catch (error) {
    console.warn(`Database query error for row ${rowNumber}:`, error.message);
    return null;
  }
}

function calculateRowAmount(matchRow, rowData, qStart, qEnd, locationCode, quarter) {
  const noOfItems = matchRow.no_of_assets || 0;
  const unitPrice = matchRow.unit_price || 0;
  const serviceDays = calculateServiceDays(matchRow.amc_from, matchRow.amc_to, qStart, qEnd);
  
  // Enhanced validation that considers cumulative consumed quantities
  const validation = validateConsumedQuantityEnhanced(
    locationCode, 
    rowData.shortTextRaw, 
    quarter, 
    rowData.consumedQty, 
    noOfItems,
    serviceDays,
    unitPrice,
    matchRow.amc_from,
    matchRow.amc_to
  );

  // Base calculation (what it would be without adjustment)
  const baseAmount = noOfItems * serviceDays * unitPrice;
  
  // Final amount with adjustment
 const finalAmount = baseAmount - validation.adjustmentAmount;

  const adjustmentAmount = finalAmount - baseAmount;
  
  let remarks = '';
  
  // Generate remarks only if there's a significant adjustment
  if (Math.abs(validation.discrepancy) > 0.01) {
    const adjustmentType = validation.discrepancy > 0 ? 'Excess' : 'Deficit';
    const discrepancyAbs = Math.abs(validation.discrepancy);
    
    remarks = `${adjustmentType} consumption detected. ` +
              `Previous quarters expected: ${validation.expectedConsumedQty.toFixed(0)}, ` +
              `Previous quarters actual: ${validation.consumedQtyFromFile}, ` +
              `Variance: ${validation.discrepancy.toFixed(0)}, ` +
              `${quarter} service days adjusted: ${serviceDays} → ${validation.adjustedServiceDays.toFixed(2)}, ` +
              `Amount adjustment: ${adjustmentAmount.toFixed(2)}`;
    
    console.log(`🔄 Previous quarters variance adjustment for ${rowData.shortTextRaw} in ${quarter}:`);
    console.log(`   Expected consumed qty (prev quarters): ${validation.expectedConsumedQty.toFixed(0)}`);
    console.log(`   Actual consumed qty (prev quarters): ${validation.consumedQtyFromFile}`);
    console.log(`   Previous quarters variance: ${validation.discrepancy.toFixed(0)}`);
    console.log(`   Current quarter service days: ${serviceDays} → ${validation.adjustedServiceDays.toFixed(2)}`);
    console.log(`   Amount: ${baseAmount.toFixed(2)} → ${finalAmount.toFixed(2)} (${adjustmentAmount.toFixed(2)})`);
  }

  return {
    noOfItems,
    serviceDays: validation.adjustedServiceDays,
    baseAmount,
    adjustmentAmount,
    amount: finalAmount,
    remarks,
    validation: {
      consumedQty: validation.consumedQtyFromFile,
      expectedQty: validation.expectedConsumedQty,
      discrepancy: validation.discrepancy,
      adjustedServiceDays: validation.adjustedServiceDays
    }
  };
}

function updateRowWithCalculations(row, colMap, calculationResult) {
  if (colMap.noOfItems) {
    row.getCell(colMap.noOfItems).value = calculationResult.noOfItems;
  }
  if (colMap.serviceDays) {
    row.getCell(colMap.serviceDays).value = calculationResult.serviceDays;
  }
  
  row.getCell(colMap.amount).value = calculationResult.amount;
  
  if (calculationResult.remarks) {
    const remarksCol = colMap.remarks || (row.worksheet.getRow(1).cellCount + 2);
    row.getCell(remarksCol).value = calculationResult.remarks;
  }
  
  row.commit();
}

function addTotalRows(uploadSheet, calcSheet, totalAmount, colMap) {
  if (totalAmount > 0) {
    const headerRow = uploadSheet.getRow(1);
    
    const uploadTotalRow = new Array(headerRow.cellCount).fill(null);
    uploadTotalRow[0] = 'TOTAL';
    uploadSheet.addRow(uploadTotalRow);

    const calcTotalRow = new Array(headerRow.cellCount).fill(null);
    calcTotalRow[0] = 'TOTAL';
    calcTotalRow[colMap.amount - 1] = totalAmount;
    calcSheet.addRow(calcTotalRow);
  }
}

async function saveFilesAndCreateZip(paths, uploadWB, calcWB, folderName) {
  const uploadPath = path.join(paths.uploadDir, `upload_${folderName}.xlsx`);
  const calcPath = path.join(paths.calcDir, `calculated_${folderName}.xlsx`);

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
  archive.directory(paths.baseDir, false);
  await archive.finalize();

  return zipFile;
}

function storeProcessedData(locationCode, lineId, quarter, officeInfo, processingResults) {
  const uniqueKey = `${locationCode}_${lineId}_${quarter}`;

  processedData.set(uniqueKey, {
    locationCode,
    actualLocationCode: officeInfo.actualLocationCode,
    lineId,
    lineNumber: officeInfo.lineNo,
    locationName: officeInfo.office ? officeInfo.office.name : officeInfo.officeName,
    quarter,
    totalAmount: processingResults.totalAmount,
    baseAmount: processingResults.baseAmount || 0, // Add base amount tracking
    adjustmentAmount: processingResults.adjustmentAmount || 0, // Add adjustment tracking
    folderName: officeInfo.folderName,
    processedAt: new Date(),
    folderPath: officeInfo.baseDir,
    totalRows: processingResults.processedRows + processingResults.skippedRows.length,
    processedRows: processingResults.processedRows,
    skippedRows: processingResults.skippedRows,
    consumedQtyValidations: processingResults.consumedQtyValidations || [] // Track validations
  });
}


function cleanupTempFiles(filePath, baseDir) {
  try {
    if (fs.existsSync(filePath)) {
      fs.unlinkSync(filePath);
    }
  } catch (error) {
    console.warn('Failed to cleanup temp files:', error.message);
  }
}

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
  } catch (error) {
    console.error('❌ Error generating location-wise report:', error);
    res.status(500).json({ error: 'Failed to generate location-wise report', details: error.message });
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
   // Clean up if exists
    if (fs.existsSync(mainFolderPath)) {
      fs.rmSync(mainFolderPath, { recursive: true, force: true });
    }

    // Create directory structure
    fs.mkdirSync(invoicesSubfolder, { recursive: true });
    fs.mkdirSync(reportsSubfolder, { recursive: true });


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