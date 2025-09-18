const express = require('express');
const router = express.Router();
const db = require('../db/connection');

// GET /dashboard/search - Search assets (AMC data)
router.get('/search', async (req, res) => {
  try {
    const q = req.query.q || '';
    const query = `
      SELECT sr_no, plant_code, location, main_line_text, service_short_text, 
             amc_from, amc_to, no_of_assets as no_of_asset, no_of_days, quantity, unit_price, total_cost
      FROM amc_data
      WHERE main_line_text LIKE ? 
      OR service_short_text LIKE ?
      OR plant_code LIKE ?
      ORDER BY sr_no
      LIMIT 100
    `;
    const values = [`%${q}%`, `%${q}%`, `%${q}%`];

    const [results] = await db.execute(query, values);
    res.json(results);
  } catch (error) {
    console.error('Error searching assets:', error);
    res.status(500).json({ message: 'Failed to fetch data', error: error.message });
  }
});

// PUT /dashboard/update/:sr_no - Update asset (AMC data)
router.put('/update/:sr_no', async (req, res) => {
  try {
    const sr_no = req.params.sr_no;
    const data = req.body;

    // Calculate derived values
    const amc_from = new Date(data.amc_from);
    const amc_to = new Date(data.amc_to);
    const no_of_days = Math.floor((amc_to - amc_from) / (1000 * 60 * 60 * 24)) + 1;
    const quantity = data.no_of_asset * no_of_days;
    const total_cost = data.no_of_asset * quantity * data.unit_price;

    const query = `
      UPDATE amc_data SET 
        location = ?, 
        main_line_text = ?, 
        service_short_text = ?, 
        amc_from = ?, 
        amc_to = ?, 
        no_of_assets = ?, 
        no_of_days = ?, 
        quantity = ?, 
        unit_price = ?, 
        total_cost = ?
      WHERE sr_no = ?
    `;

    const values = [
      data.location,
      data.main_line_text,
      data.service_short_text,
      data.amc_from,
      data.amc_to,
      data.no_of_asset,
      no_of_days,
      quantity,
      data.unit_price,
      total_cost,
      sr_no
    ];

    const [result] = await db.execute(query, values);
    
    if (result.affectedRows > 0) {
      res.json({ 
        message: 'Asset updated successfully',
        updatedData: {
          no_of_days,
          quantity,
          total_cost
        }
      });
    } else {
      res.status(404).json({ message: 'Asset not found' });
    }
  } catch (error) {
    console.error('Error updating asset:', error);
    res.status(500).json({ message: 'Error updating asset', error: error.message });
  }
});

// GET /dashboard/assets - Get all assets with pagination (AMC data)
router.get('/assets', async (req, res) => {
  try {
    const page = parseInt(req.query.page) || 1;
    const limit = parseInt(req.query.limit) || 50;
    const offset = (page - 1) * limit;

    const query = `
      SELECT sr_no, plant_code, location, main_line_text, service_short_text, 
             amc_from, amc_to, no_of_assets as no_of_asset, no_of_days, quantity, unit_price, total_cost
      FROM amc_data
      ORDER BY sr_no
      LIMIT ? OFFSET ?
    `;

    const countQuery = `SELECT COUNT(*) as total FROM amc_data`;

    const [results] = await db.execute(query, [limit, offset]);
    const [countResult] = await db.execute(countQuery);

    res.json({
      data: results,
      pagination: {
        page,
        limit,
        total: countResult[0].total,
        totalPages: Math.ceil(countResult[0].total / limit)
      }
    });
  } catch (error) {
    console.error('Error fetching assets:', error);
    res.status(500).json({ message: 'Failed to fetch assets', error: error.message });
  }
});

// DELETE /dashboard/delete/:sr_no - Delete asset (AMC data)
router.delete('/delete/:sr_no', async (req, res) => {
  try {
    const sr_no = req.params.sr_no;
    const query = `DELETE FROM amc_data WHERE sr_no = ?`;
    
    const [result] = await db.execute(query, [sr_no]);
    
    if (result.affectedRows > 0) {
      res.json({ message: 'Asset deleted successfully' });
    } else {
      res.status(404).json({ message: 'Asset not found' });
    }
  } catch (error) {
    console.error('Error deleting asset:', error);
    res.status(500).json({ message: 'Error deleting asset', error: error.message });
  }
});

// POST /dashboard/add - Add new asset (AMC data)
router.post('/add', async (req, res) => {
  try {
    const data = req.body;
    
    // Calculate derived values
    const amc_from = new Date(data.amc_from);
    const amc_to = new Date(data.amc_to);
    const no_of_days = Math.floor((amc_to - amc_from) / (1000 * 60 * 60 * 24)) + 1;
    const quantity = data.no_of_asset * no_of_days;
    const total_cost = data.no_of_asset * quantity * data.unit_price;

    const query = `
      INSERT INTO amc_data (
        plant_code, location, main_line_text, service_short_text, 
        amc_from, amc_to, no_of_assets, no_of_days, quantity, unit_price, total_cost
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `;

    const values = [
      data.plant_code,
      data.location,
      data.main_line_text,
      data.service_short_text,
      data.amc_from,
      data.amc_to,
      data.no_of_asset,
      no_of_days,
      quantity,
      data.unit_price,
      total_cost
    ];

    const [result] = await db.execute(query, values);
    
    res.json({ 
      message: 'Asset added successfully',
      assetId: result.insertId,
      calculatedData: {
        no_of_days,
        quantity,
        total_cost
      }
    });
  } catch (error) {
    console.error('Error adding asset:', error);
    res.status(500).json({ message: 'Error adding asset', error: error.message });
  }
});

module.exports = router;