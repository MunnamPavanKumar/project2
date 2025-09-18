const express = require('express');
const mysql = require('mysql2/promise');
const cors = require('cors');
const router = express.Router();
require('dotenv').config();
const authenticateToken = require('../middleware/authMiddleware'); 
const pool = require('../db/connection');
// Middleware
router.use(cors());
router.use(express.json());
router.use(authenticateToken);



// Initialize database and table
async function initDatabase() {
  try {
    const connection = await pool.getConnection();
    connection.release();
    console.log('Database and tables initialized successfully');
  } catch (error) {
    console.error('Database initialization error:', error);
    process.exit(1);
  }
}

// Helper function to log audit trail (disabled)
async function logAuditTrail(assetId, action, oldValues = null, newValues = null) {
  console.log(`Audit log (disabled): ${action} for asset ${assetId}`);
}

// Routes
router.get('/assets', async (req, res) => {
  try {
    const { q, limit, offset } = req.query;
    let query = 'SELECT * FROM amc_data';
    let params = [];

    if (q && q.trim() !== '') {
      query += ' WHERE sr_no LIKE ? OR main_line_short_text LIKE ? OR plant_code LIKE ?';
      params = [`%${q}%`, `%${q}%`, `%${q}%`];
    }

    query += ' ORDER BY sr_no DESC';

    if (limit) {
      query += ' LIMIT ?';
      params.push(parseInt(limit));
      if (offset) {
        query += ' OFFSET ?';
        params.push(parseInt(offset));
      }
    }

    const [rows] = await pool.execute(query, params);
    res.json(rows);
  } catch (error) {
    console.error('Error fetching assets:', error);
    res.status(500).json({ error: 'Internal server error' });
  }
});

router.get('/assets/count', async (req, res) => {
  try {
    const { q } = req.query;
    let query = 'SELECT COUNT(*) as total FROM amc_data';
    let params = [];

    if (q && q.trim() !== '') {
      query += ' WHERE sr_no LIKE ? OR main_line_short_text LIKE ? OR plant_code LIKE ?';
      params = [`%${q}%`, `%${q}%`, `%${q}%`];
    }

    const [rows] = await pool.execute(query, params);
    res.json({ total: rows[0].total });
  } catch (error) {
    console.error('Error counting assets:', error);
    res.status(500).json({ error: 'Internal server error' });
  }
});

router.get('/assets/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const [rows] = await pool.execute('SELECT * FROM amc_data WHERE sr_no = ?', [id]);

    if (rows.length === 0) {
      return res.status(404).json({ error: 'Asset not found' });
    }

    res.json(rows[0]);
  } catch (error) {
    console.error('Error fetching asset:', error);
    res.status(500).json({ error: 'Internal server error' });
  }
});

router.post('/assets', async (req, res) => {
  try {
    const {
      main_line_short_text,
      plant_code,
      amc_from,
      amc_to,
      no_of_asset,
      quantity,
      unit_price
    } = req.body;

    if (!main_line_short_text || !plant_code || !amc_from || !amc_to) {
      return res.status(400).json({ error: 'Missing required fields' });
    }

    const fromDate = new Date(amc_from);
    const toDate = new Date(amc_to);
    if (fromDate >= toDate) {
      return res.status(400).json({ error: 'AMC From date must be before AMC To date' });
    }

    const numericFields = {
      no_of_asset: no_of_asset || 1,
      quantity: quantity || 1,
      unit_price: unit_price || 0
    };

    for (const [field, value] of Object.entries(numericFields)) {
      if (isNaN(value) || value < 0) {
        return res.status(400).json({ error: `Invalid value for ${field}` });
      }
    }

    const [result] = await pool.execute(
      `INSERT INTO amc_data (main_line_short_text, plant_code, amc_from, amc_to, no_of_asset, quantity, unit_price) 
       VALUES (?, ?, ?, ?, ?, ?, ?)`,
      [main_line_short_text, plant_code, amc_from, amc_to, numericFields.no_of_asset, numericFields.quantity, numericFields.unit_price]
    );

    const [newAsset] = await pool.execute('SELECT * FROM amc_data WHERE sr_no = ?', [result.insertId]);
    await logAuditTrail(result.insertId, 'CREATE', null, newAsset[0]);

    res.status(201).json(newAsset[0]);
  } catch (error) {
    console.error('Error creating asset:', error);
    res.status(500).json({ error: 'Internal server error' });
  }
});

router.put('/assets/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const updates = req.body;

    const [existing] = await pool.execute('SELECT * FROM amc_data WHERE sr_no = ?', [id]);
    if (existing.length === 0) {
      return res.status(404).json({ error: 'Asset not found' });
    }

    const oldValues = existing[0];
    const updateFields = [];
    const updateValues = [];

    const allowedFields = ['service_short_text', 'plant_code', 'amc_from', 'amc_to', 'no_of_asset', 'quantity', 'unit_price'];

    for (const field of allowedFields) {
      if (updates[field] !== undefined) {
        if (['no_of_asset', 'quantity', 'unit_price'].includes(field)) {
          const numValue = parseFloat(updates[field]);
          if (isNaN(numValue) || numValue < 0) {
            return res.status(400).json({ error: `Invalid value for ${field}` });
          }
          updateFields.push(`${field} = ?`);
          updateValues.push(numValue);
        } else if (field === 'amc_from' || field === 'amc_to') {
          const date = new Date(updates[field]);
          if (isNaN(date.getTime())) {
            return res.status(400).json({ error: `Invalid date for ${field}` });
          }
          updateFields.push(`${field} = ?`);
          updateValues.push(updates[field]);
        } else {
          updateFields.push(`${field} = ?`);
          updateValues.push(updates[field]);
        }
      }
    }

    if (updateFields.length === 0) {
      return res.status(400).json({ error: 'No valid fields to update' });
    }

    let fromDate = oldValues.amc_from;
    let toDate = oldValues.amc_to;

    if (updates.amc_from) fromDate = new Date(updates.amc_from);
    if (updates.amc_to) toDate = new Date(updates.amc_to);

    if (fromDate >= toDate) {
      return res.status(400).json({ error: 'AMC From date must be before AMC To date' });
    }

    updateValues.push(id);

    await pool.execute(
      `UPDATE amc_data SET ${updateFields.join(', ')} WHERE sr_no = ?`,
      updateValues
    );

    const [updatedAsset] = await pool.execute('SELECT * FROM amc_data WHERE sr_no = ?', [id]);
    await logAuditTrail(id, 'UPDATE', oldValues, updatedAsset[0]);

    console.log(`Asset ${id} updated:`, updates);
    res.json(updatedAsset[0]);
  } catch (error) {
    console.error('Error updating asset:', error);
    res.status(500).json({ error: 'Internal server error' });
  }
});

router.delete('/assets/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const [existing] = await pool.execute('SELECT * FROM amc_data WHERE sr_no = ?', [id]);
    if (existing.length === 0) {
      return res.status(404).json({ error: 'Asset not found' });
    }

    const oldValues = existing[0];
    await pool.execute('DELETE FROM amc_data WHERE sr_no = ?', [id]);
    await logAuditTrail(id, 'DELETE', oldValues, null);

    res.json({ message: 'Asset deleted successfully' });
  } catch (error) {
    console.error('Error deleting asset:', error);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// Batch update, import, summary report, expiring, and health check remain the same
// Just ensure all SQL references are to `amc_data` instead of `assets`

// [rest of the file remains same, including startServer()]

// Add batch update endpoint for multiple assets
router.put('/assets/batch', async (req, res) => {
  try {
    const { updates } = req.body; // Array of {id, data} objects
    
    if (!Array.isArray(updates) || updates.length === 0) {
      return res.status(400).json({ error: 'Invalid batch update data' });
    }
    
    const results = [];
    
    // Start transaction
    const connection = await pool.getConnection();
    await connection.beginTransaction();
    
    try {
      for (const update of updates) {
        const { id, data } = update;
        
        // Check if asset exists
        const [existing] = await connection.execute('SELECT * FROM amc_data WHERE sr_no = ?', [id]);
        if (existing.length === 0) {
          throw new Error(`Asset ${id} not found`);
        }
        
        const oldValues = existing[0];
        
        // Build update query
        const updateFields = [];
        const updateValues = [];
        const allowedFields = ['main_line_short_text', 'plant_code', 'amc_from', 'amc_to', 'no_of_asset', 'quantity', 'unit_price'];
        
        for (const field of allowedFields) {
          if (data[field] !== undefined) {
            if (['no_of_asset', 'quantity', 'unit_price'].includes(field)) {
              const numValue = parseFloat(data[field]);
              if (isNaN(numValue) || numValue < 0) {
                throw new Error(`Invalid value for ${field} in asset ${id}`);
              }
              updateFields.push(`${field} = ?`);
              updateValues.push(numValue);
            } else {
              updateFields.push(`${field} = ?`);
              updateValues.push(data[field]);
            }
          }
        }
        
        if (updateFields.length > 0) {
          updateValues.push(id);
          await connection.execute(
            `UPDATE amc_data SET ${updateFields.join(', ')} WHERE sr_no = ?`,
            updateValues
          );
          
          const [updatedAsset] = await connection.execute('SELECT * FROM amc_data WHERE sr_no = ?', [id]);
          
          // Log audit trail
          await logAuditTrail(id, 'BATCH_UPDATE', oldValues, updatedAsset[0]);
        }
        
        results.push({ id, status: 'updated' });
      }
      
      await connection.commit();
      res.json({ success: true, results });
    } catch (error) {
      await connection.rollback();
      throw error;
    } finally {
      connection.release();
    }
  } catch (error) {
    console.error('Error in batch update:', error);
    res.status(500).json({ error: 'Batch update failed: ' + error.message });
  }
});

// GET /api/assets/:id/audit - Get audit trail for an asset (disabled)
router.get('/assets/:id/audit', async (req, res) => {
  try {
    const { id } = req.params;
    // Return empty array since audit logging is disabled
    res.json([]);
  } catch (error) {
    console.error('Error fetching audit trail:', error);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// GET /api/reports/summary - Get summary report
router.get('/reports/summary', async (req, res) => {
  try {
    const [totalAssets] = await pool.execute('SELECT COUNT(*) as total FROM amc_data');
    const [totalValue] = await pool.execute('SELECT SUM(quantity * unit_price) as total_value FROM amc_data');
    const [plantCounts] = await pool.execute('SELECT plant_code, COUNT(*) as count FROM amc_data GROUP BY plant_code');
    const [expiringAssets] = await pool.execute(
      'SELECT COUNT(*) as expiring FROM amc_data WHERE amc_to <= DATE_ADD(CURDATE(), INTERVAL 30 DAY)'
    );
    
    res.json({
      total_assets: totalAssets[0].total,
      total_value: totalValue[0].total_value || 0,
      plant_distribution: plantCounts,
      expiring_in_30_days: expiringAssets[0].expiring
    });
  } catch (error) {
    console.error('Error generating summary report:', error);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// GET /api/reports/expiring - Get assets expiring soon
router.get('/reports/expiring', async (req, res) => {
  try {
    const { days = 30 } = req.query;
    const [rows] = await pool.execute(
      'SELECT * FROM amc_data WHERE amc_to <= DATE_ADD(CURDATE(), INTERVAL ? DAY) ORDER BY amc_to ASC',
      [parseInt(days)]
    );
    res.json(rows);
  } catch (error) {
    console.error('Error fetching expiring assets:', error);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// POST /api/assets/import - Import assets from CSV/JSON
router.post('/assets/import', async (req, res) => {
  try {
    const { assets } = req.body;
    
    if (!Array.isArray(assets) || assets.length === 0) {
      return res.status(400).json({ error: 'Invalid import data' });
    }
    
    const results = [];
    const connection = await pool.getConnection();
    await connection.beginTransaction();
    
    try {
      for (const asset of assets) {
        const {
          main_line_short_text,
          plant_code,
          amc_from,
          amc_to,
          no_of_asset = 1,
          quantity = 1,
          unit_price = 0
        } = asset;
        
        // Validate required fields
        if (!main_line_short_text || !plant_code || !amc_from || !amc_to) {
          results.push({ asset, status: 'failed', error: 'Missing required fields' });
          continue;
        }
        
        try {
          const [result] = await connection.execute(
            `INSERT INTO amc_data (main_line_short_text, plant_code, amc_from, amc_to, no_of_asset, quantity, unit_price) 
             VALUES (?, ?, ?, ?, ?, ?, ?)`,
            [main_line_short_text, plant_code, amc_from, amc_to, no_of_asset, quantity, unit_price]
          );
          
          results.push({ asset, status: 'success', id: result.insertId });
        } catch (error) {
          results.push({ asset, status: 'failed', error: error.message });
        }
      }
      
      await connection.commit();
      res.json({ results });
    } catch (error) {
      await connection.rollback();
      throw error;
    } finally {
      connection.release();
    }
  } catch (error) {
    console.error('Error importing assets:', error);
    res.status(500).json({ error: 'Import failed: ' + error.message });
  }
});

// Health check endpoint
router.get('/health', (req, res) => {
  res.json({ status: 'OK', timestamp: new Date().toISOString() });
});

// Error handling middleware
router.use((err, req, res, next) => {
  console.error(err.stack);
  res.status(500).json({ error: 'Something went wrong!' });
});

// 404 handler
router.use((req, res) => {
  res.status(404).json({ error: 'Route not found' });
});



// Graceful shutdown
process.on('SIGINT', async () => {

  await pool.end();
  process.exit(0);
});


process.on('SIGTERM', async () => {
  console.log('\nShutting down server...');
  await pool.end();
  process.exit(0);
});

module.exports=router;