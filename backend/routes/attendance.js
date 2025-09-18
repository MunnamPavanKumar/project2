const express = require('express');
const mysql = require('mysql2/promise');
const cors = require('cors');
const ExcelJS = require('exceljs');
const path = require('path');

const router = express.Router();
const PORT = 3000;

// Middleware
router.use(cors());
router.use(express.json());

const pool = require('../db/connection');



// Constants
const PENALTY_PER_DAY = 1498;
const GST_RATE = 0.18;



// Helper function to get quarter info - Updated to match your requirements
function getQuarterInfo() {
    const quarters = [];
    const currentDate = new Date();
    const currentMonth = currentDate.getMonth() + 1; // 1-12
    const currentYear = currentDate.getFullYear();
    
    // Define quarters starting from Q1 (Jun 2024 - Aug 2024)
    const quarterDefinitions = [
        { key: 'q1', name: 'Quarter 1 (Jun-Aug 2024)', startMonth: 6, endMonth: 8, year: 2024 },
        { key: 'q2', name: 'Quarter 2 (Sep-Nov 2024)', startMonth: 9, endMonth: 11, year: 2024 },
        { key: 'q3', name: 'Quarter 3 (Dec 2024-Feb 2025)', startMonth: 12, endMonth: 2, year: 2024, crossYear: true, endYear: 2025 },
        { key: 'q4', name: 'Quarter 4 (Mar-May 2025)', startMonth: 3, endMonth: 5, year: 2025 },
        { key: 'q5', name: 'Quarter 5 (Jun-Aug 2025)', startMonth: 6, endMonth: 8, year: 2025 },
        { key: 'q6', name: 'Quarter 6 (Sep-Nov 2025)', startMonth: 9, endMonth: 11, year: 2025 },
        { key: 'q7', name: 'Quarter 7 (Dec 2025-Feb 2026)', startMonth: 12, endMonth: 2, year: 2025, crossYear: true, endYear: 2026 },
        { key: 'q8', name: 'Quarter 8 (Mar-May 2026)', startMonth: 3, endMonth: 5, year: 2026 },
        { key: 'q9', name: 'Quarter 9 (Jun-Aug 2026)', startMonth: 6, endMonth: 8, year: 2026 },
        { key: 'q10', name: 'Quarter 10 (Sep-Nov 2026)', startMonth: 9, endMonth: 11, year: 2026 },
        { key: 'q11', name: 'Quarter 11 (Dec 2026-Feb 2027)', startMonth: 12, endMonth: 2, year: 2026, crossYear: true, endYear: 2027 },
        { key: 'q12', name: 'Quarter 12 (Mar-May 2027)', startMonth: 3, endMonth: 5, year: 2027 }
    ];
    
    // Function to check if a quarter has started based on current date
    function hasQuarterStarted(quarter) {
        if (quarter.crossYear) {
            // For cross-year quarters (Dec-Feb), check if we're past the start month of the first year
            // or if we're in the second year and past the start (January)
            return (currentYear > quarter.year) || 
                   (currentYear === quarter.year && currentMonth >= quarter.startMonth) ||
                   (currentYear === quarter.endYear && currentMonth >= 1);
        } else {
            // For normal quarters, check if we're in the same year and past the start month
            return (currentYear > quarter.year) || 
                   (currentYear === quarter.year && currentMonth >= quarter.startMonth);
        }
    }
    
    // Only return quarters that have started
    return quarterDefinitions.filter(quarter => hasQuarterStarted(quarter));
}

// Helper function to get months in a quarter
function getMonthsInQuarter(quarter) {
    const quarterInfo = getQuarterInfo().find(q => q.key === quarter);
    if (!quarterInfo) return [];
    
    const months = [];
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                       'July', 'August', 'September', 'October', 'November', 'December'];
    
    if (quarterInfo.crossYear) {
        // Handle Dec-Feb quarter that crosses year boundary
        // Add December of start year
        months.push({
            key: `${quarterInfo.year}-12`,
            name: `${monthNames[11]} ${quarterInfo.year}`,
            month: 12,
            year: quarterInfo.year
        });
        
        // Add January and February of end year
        for (let month = 1; month <= quarterInfo.endMonth; month++) {
            months.push({
                key: `${quarterInfo.endYear}-${month.toString().padStart(2, '0')}`,
                name: `${monthNames[month-1]} ${quarterInfo.endYear}`,
                month: month,
                year: quarterInfo.endYear
            });
        }
    } else {
        // Normal quarters within same year
        for (let month = quarterInfo.startMonth; month <= quarterInfo.endMonth; month++) {
            months.push({
                key: `${quarterInfo.year}-${month.toString().padStart(2, '0')}`,
                name: `${monthNames[month-1]} ${quarterInfo.year}`,
                month: month,
                year: quarterInfo.year
            });
        }
    }
    
    return months;
}

// Helper function to get days in a month
function getDaysInMonth(year, month) {
    const days = [];
    const daysCount = new Date(year, month, 0).getDate();
    
    for (let day = 1; day <= daysCount; day++) {
        const date = new Date(year, month - 1, day);
        days.push({
            date: date.toISOString().split('T')[0],
            dayNumber: day,
            dayName: date.toLocaleDateString('en-US', { weekday: 'short' })
        });
    }
    
    return days;
}

// Initialize database tables
async function initializeDatabase() {
    try {
        const connection = await pool.getConnection();
        
        // Create employees table if not exists
        await connection.execute(`
            CREATE TABLE IF NOT EXISTS employees (
               id VARCHAR(20) PRIMARY KEY,
                name VARCHAR(255) NOT NULL,
                designation VARCHAR(255),
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        `);
        
        // Create attendance table if not exists
        await connection.execute(`
            CREATE TABLE IF NOT EXISTS attendance (
               id VARCHAR(20) PRIMARY KEY,
                employee_id VARCHAR(20) NOT NULL,
                date DATE NOT NULL,
                status ENUM('present', 'absent', 'awor') DEFAULT 'present',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE,
                UNIQUE KEY unique_employee_date (employee_id, date)
            )
        `);
        
        // Insert CIPL employees if table is empty
        const [rows] = await connection.execute('SELECT COUNT(*) as count FROM employees');
        if (rows[0].count === 0) {
            const ciplEmployees = [
                ['A005082', 'VINOTHKUMAR K', 'System Administrator(TL)'],
                ['A005083', 'RAGUL PACKIRISAMY', 'System Administrator'],
                ['A005184', 'PRAKASH A', 'Senior System Administrator'],
                ['A005086', 'Vigneshwaran', 'Network Engineer'],
                ['A005185', 'M. Muthukrishnan', 'System Administrator'],
                ['A005087', 'A.GOKULAN', 'System Administrator'],
                ['A005089', 'Vignesh kumar D', 'System Administrator'],
                ['A005186', 'SADISH G', 'Network Engineer'],
                ['A005249', 'POUNRAJ K', 'Desktop support Engineer'],
                ['A005288', 'Prithiyuman A', 'Desktop support Engineer'],
                ['A005377', 'Hari Chakkaravarthi Raj', 'Desktop support Engineer'],
                ['A005808', 'Vijayalayan', 'Desktop support Engineer'],
                ['A005923', 'Magesh kumar', 'Desktop support Engineer'],
                ['A005232', 'Kalaivendhan', 'Server Engineer'],
                ['A005231', 'Sakthi Saravanan.S', 'Server Engineer'],
                ['A005233', 'Pushparaj', 'Server Engineer'],
                ['A005234', 'Kalainesan P', 'Server Engineer'],
                ['A005229', 'JAIVANTH ASIRVATHAM R', 'Network Administrator'],
                ['A005230', 'ARUL BOOPATHY B', 'Network Administrator'],
                ['A005228', 'SUTHARSAN K', 'Network Administrator'],
                ['A005898', 'PRINCE', 'Network Administrator']
            ];
            
            for (const [empId, name, designation] of ciplEmployees) {
                await connection.execute(
                    'INSERT INTO employees (id, name, designation) VALUES (?, ?, ?)',
                    [empId, name, designation]
                );
            }
        }
        
        connection.release();
        console.log('Database initialized successfully');
    } catch (error) {
        console.error('Database initialization error:', error);
    }
}

// API Routes
initializeDatabase();
// Get all available quarters
router.get('/quarters', (req, res) => {
    try {
        const quarters = getQuarterInfo();
        res.json(quarters);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Get months in a quarter
router.get('/quarters/:quarter/months', (req, res) => {
    try {
        const { quarter } = req.params;
        const months = getMonthsInQuarter(quarter);
        res.json(months);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Get days in a month
router.get('/months/:monthKey/days', (req, res) => {
    try {
        const { monthKey } = req.params;
        const [year, month] = monthKey.split('-').map(Number);
        const days = getDaysInMonth(year, month);
        res.json(days);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Get all employees
router.get('/employees', async (req, res) => {
    try {
        const connection = await pool.getConnection();
        const [rows] = await connection.execute(
            'SELECT id, name, designation FROM employees ORDER BY name'
        );
        connection.release();
        res.json(rows);
    } catch (error) {
        console.error('Error fetching employees:', error);
        res.status(500).json({ error: error.message });
    }
});

// Get employees for a specific date
router.get('/employees/date/:date', async (req, res) => {
    try {
        const { date } = req.params;
        const connection = await pool.getConnection();
        
        const [rows] = await connection.execute(`
            SELECT 
                e.id,
                e.name,
                e.designation,
                COALESCE(a.status, 'present') as status
            FROM employees e
            LEFT JOIN attendance a ON e.id = a.employee_id AND a.date = ?
            ORDER BY e.name
        `, [date]);
        
        connection.release();
        res.json(rows);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

router.post('/attendance', async (req, res) => {
    try {
        const { employeeId, date, status } = req.body;
        
        // Validate input
        if (!employeeId || !date || !status) {
            return res.status(400).json({ error: 'Missing required fields: employeeId, date, status' });
        }
        
        if (!['present', 'absent', 'awor'].includes(status)) {
            return res.status(400).json({ error: 'Invalid status. Must be: present, absent, or awor' });
        }
        
        const connection = await pool.getConnection();
        
        try {
            // First check if record exists
            const [existing] = await connection.execute(
                'SELECT id FROM attendance WHERE employee_id = ? AND date = ?',
                [employeeId, date]
            );
            
            if (existing.length > 0) {
                // Update existing record
                await connection.execute(
                    'UPDATE attendance SET status = ?, updated_at = CURRENT_TIMESTAMP WHERE employee_id = ? AND date = ?',
                    [status, employeeId, date]
                );
            } else {
                // Insert new record with generated ID
                const attendanceId = `${employeeId}_${date}`;
                console.log(`Generated attendance ID: ${attendanceId}`);
                await connection.execute(
                    'INSERT INTO attendance (id, employee_id, date, status) VALUES (?, ?, ?, ?)',
                    [attendanceId, employeeId, date, status]
                );
            }
        } finally {
            connection.release();
        }
        
        res.json({ 
            success: true, 
            message: 'Attendance saved successfully',
            data: { employeeId, date, status }
        });
        
    } catch (error) {
        console.error('Error saving attendance:', error);
        res.status(500).json({ 
            error: 'Failed to save attendance',
            details: error.message 
        });
    }
});
// Get attendance for a specific employee
router.get('/attendance/:employeeId', async (req, res) => {
    try {
        const { employeeId } = req.params;
        const { quarter } = req.query;
        
        let query = 'SELECT date, status FROM attendance WHERE employee_id = ?';
        let params = [employeeId];
        
        if (quarter) {
            const quarterInfo = getQuarterInfo().find(q => q.key === quarter);
            if (quarterInfo) {
                if (quarterInfo.crossYear) {
                    // Handle cross-year quarter (Dec-Feb)
                    query += ' AND ((YEAR(date) = ? AND MONTH(date) = ?) OR (YEAR(date) = ? AND MONTH(date) BETWEEN ? AND ?))';
                    params.push(quarterInfo.year, quarterInfo.startMonth, quarterInfo.endYear, 1, quarterInfo.endMonth);
                } else {
                    // Normal quarter within same year
                    query += ' AND YEAR(date) = ? AND MONTH(date) BETWEEN ? AND ?';
                    params.push(quarterInfo.year, quarterInfo.startMonth, quarterInfo.endMonth);
                }
            }
        }
        
        query += ' ORDER BY date';
        
        const connection = await pool.getConnection();
        const [rows] = await connection.execute(query, params);
        connection.release();
        
        res.json(rows);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Get attendance data for a specific date range or quarter
router.get('/attendance/range/:quarter', async (req, res) => {
    try {
        const { quarter } = req.params;
        const quarterInfo = getQuarterInfo().find(q => q.key === quarter);
        
        if (!quarterInfo) {
            return res.status(404).json({ error: 'Quarter not found' });
        }
        
        const connection = await pool.getConnection();
        let query, params;
        
        if (quarterInfo.crossYear) {
            // Handle cross-year quarter (Dec-Feb)
            query = `
                SELECT 
                    employee_id,
                    date,
                    status
                FROM attendance 
                WHERE ((YEAR(date) = ? AND MONTH(date) = ?) OR (YEAR(date) = ? AND MONTH(date) BETWEEN ? AND ?))
                ORDER BY date, employee_id
            `;
            params = [quarterInfo.year, quarterInfo.startMonth, quarterInfo.endYear, 1, quarterInfo.endMonth];
        } else {
            // Normal quarter within same year
            query = `
                SELECT 
                    employee_id,
                    date,
                    status
                FROM attendance 
                WHERE YEAR(date) = ? AND MONTH(date) BETWEEN ? AND ?
                ORDER BY date, employee_id
            `;
            params = [quarterInfo.year, quarterInfo.startMonth, quarterInfo.endMonth];
        }
        
        const [rows] = await connection.execute(query, params);
        connection.release();
        
        // Transform the data into a more frontend-friendly format
        const attendanceData = {};
        rows.forEach(row => {
            const dateStr = row.date.toISOString().split('T')[0];
            if (!attendanceData[dateStr]) {
                attendanceData[dateStr] = {};
            }
            attendanceData[dateStr][row.employee_id] = row.status;
        });
        
        res.json(attendanceData);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Get attendance summary for a quarter
router.get('/attendance/summary/:quarter', async (req, res) => {
    try {
        const { quarter } = req.params;
        const quarterInfo = getQuarterInfo().find(q => q.key === quarter);
        
        if (!quarterInfo) {
            return res.status(404).json({ error: 'Quarter not found' });
        }
        
        const connection = await pool.getConnection();
        let query, params;
        
        if (quarterInfo.crossYear) {
            // Handle cross-year quarter (Dec-Feb)
            query = `
                SELECT 
                    e.id,
                    e.name,
                    e.designation,
                    COUNT(CASE WHEN a.status = 'present' THEN 1 END) as present_days,
                    COUNT(CASE WHEN a.status = 'absent' THEN 1 END) as absent_days,
                    COUNT(CASE WHEN a.status = 'awor' THEN 1 END) as awor_days,
                    COUNT(CASE WHEN a.status = 'awor' THEN 1 END) * ? as penalty_amount,
                    COUNT(CASE WHEN a.status = 'awor' THEN 1 END) * ? * (1 + ?) as penalty_with_gst
                FROM employees e
                LEFT JOIN attendance a ON e.id = a.employee_id 
                    AND ((YEAR(a.date) = ? AND MONTH(a.date) = ?) OR (YEAR(a.date) = ? AND MONTH(a.date) BETWEEN ? AND ?))
                GROUP BY e.id, e.name, e.designation
                ORDER BY e.name
            `;
            params = [PENALTY_PER_DAY, PENALTY_PER_DAY, GST_RATE, quarterInfo.year, quarterInfo.startMonth, quarterInfo.endYear, 1, quarterInfo.endMonth];
        } else {
            // Normal quarter within same year
            query = `
                SELECT 
                    e.id,
                    e.name,
                    e.designation,
                    COUNT(CASE WHEN a.status = 'present' THEN 1 END) as present_days,
                    COUNT(CASE WHEN a.status = 'absent' THEN 1 END) as absent_days,
                    COUNT(CASE WHEN a.status = 'awor' THEN 1 END) as awor_days,
                    COUNT(CASE WHEN a.status = 'awor' THEN 1 END) * ? as penalty_amount,
                    COUNT(CASE WHEN a.status = 'awor' THEN 1 END) * ? * (1 + ?) as penalty_with_gst
                FROM employees e
                LEFT JOIN attendance a ON e.id = a.employee_id 
                    AND YEAR(a.date) = ? 
                    AND MONTH(a.date) BETWEEN ? AND ?
                GROUP BY e.id, e.name, e.designation
                ORDER BY e.name
            `;
            params = [PENALTY_PER_DAY, PENALTY_PER_DAY, GST_RATE, quarterInfo.year, quarterInfo.startMonth, quarterInfo.endMonth];
        }
        
        const [rows] = await connection.execute(query, params);
        connection.release();
        
        res.json(rows);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Export penalty report - Updated to show only total days and penalties
router.get('/export/penalties/:quarter', async (req, res) => {
    try {
        const { quarter } = req.params;
        const quarterInfo = getQuarterInfo().find(q => q.key === quarter);
        
        if (!quarterInfo) {
            return res.status(404).json({ error: 'Quarter not found' });
        }
        
        const connection = await pool.getConnection();
        let query, params;
        
        if (quarterInfo.crossYear) {
            // Handle cross-year quarter (Dec-Feb)
            query = `
                SELECT 
                    e.id,
                    e.name,
                    e.designation,
                    COUNT(CASE WHEN a.status = 'awor' THEN 1 END) as awor_days,
                    COUNT(CASE WHEN a.status = 'awor' THEN 1 END) * ? as penalty_amount,
                    COUNT(CASE WHEN a.status = 'awor' THEN 1 END) * ? * (1 + ?) as penalty_with_gst
                FROM employees e
                LEFT JOIN attendance a ON e.id = a.employee_id 
                    AND ((YEAR(a.date) = ? AND MONTH(a.date) = ?) OR (YEAR(a.date) = ? AND MONTH(a.date) BETWEEN ? AND ?))
                GROUP BY e.id, e.name, e.designation
                HAVING awor_days > 0
                ORDER BY e.name
            `;
            params = [PENALTY_PER_DAY, PENALTY_PER_DAY, GST_RATE, quarterInfo.year, quarterInfo.startMonth, quarterInfo.endYear, 1, quarterInfo.endMonth];
        } else {
            // Normal quarter within same year
            query = `
                SELECT 
                    e.id,
                    e.name,
                    e.designation,
                    COUNT(CASE WHEN a.status = 'awor' THEN 1 END) as awor_days,
                    COUNT(CASE WHEN a.status = 'awor' THEN 1 END) * ? as penalty_amount,
                    COUNT(CASE WHEN a.status = 'awor' THEN 1 END) * ? * (1 + ?) as penalty_with_gst
                FROM employees e
                LEFT JOIN attendance a ON e.id = a.employee_id 
                    AND YEAR(a.date) = ? 
                    AND MONTH(a.date) BETWEEN ? AND ?
                GROUP BY e.id, e.name, e.designation
                HAVING awor_days > 0
                ORDER BY e.name
            `;
            params = [PENALTY_PER_DAY, PENALTY_PER_DAY, GST_RATE, quarterInfo.year, quarterInfo.startMonth, quarterInfo.endMonth];
        }
        
        const [rows] = await connection.execute(query, params);
        connection.release();
        
        // Create Excel workbook
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Penalty Report');
        
        // Add headers - Updated to show only essential columns
        worksheet.columns = [
            { header: 'Employee ID', key: 'id', width: 15 },
            { header: 'Employee Name', key: 'name', width: 25 },
            { header: 'Designation', key: 'designation', width: 25 },
            { header: 'Total AWOR Days', key: 'awor_days', width: 18 },
            { header: 'Total Penalty (₹)', key: 'penalty_amount', width: 20 },
            { header: 'Total Penalty with GST (₹)', key: 'penalty_with_gst', width: 22 }
        ];
        
        // Add data rows
        rows.forEach(row => {
            worksheet.addRow({
                id: row.id,
                name: row.name,
                designation: row.designation,
                awor_days: row.awor_days,
                penalty_amount: row.penalty_amount,
                penalty_with_gst: parseFloat(row.penalty_with_gst).toFixed(2)
            });
        });
        
        // Style the header row
        worksheet.getRow(1).eachCell((cell) => {
            cell.font = { bold: true };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE0E0E0' }
            };
        });
        
        // Add totals row
        const totalAworDays = rows.reduce((sum, row) => sum + row.awor_days, 0);
        const totalPenalty = rows.reduce((sum, row) => sum + row.penalty_amount, 0);
        const totalPenaltyWithGst = rows.reduce((sum, row) => sum + parseFloat(row.penalty_with_gst), 0);
        
        worksheet.addRow({});
        const totalRow = worksheet.addRow({
            id: '',
            name: '',
            designation: 'TOTAL',
            awor_days: totalAworDays,
            penalty_amount: totalPenalty,
            penalty_with_gst: totalPenaltyWithGst.toFixed(2)
        });
        
        totalRow.eachCell((cell) => {
            cell.font = { bold: true };
        });
        
        // Set response headers
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=penalty-report-${quarter}.xlsx`);
        
        // Send the Excel file
        await workbook.xlsx.write(res);
        res.end();
        
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});




module.exports = router;