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

module.exports = router;