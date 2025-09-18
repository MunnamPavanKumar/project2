const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const processRoutes = require('./routes/process');
const dashboardRoutes = require('./routes/dashboard');
const attendanceRoutes = require('./routes/attendance');
const authenticateToken = require('./middleware/authMiddleware');
const authRoutes = require('./routes/auth'); // 👈 New auth route

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Ensure folders exist
const dirs = ['uploads', 'output', 'downloads', 'temp'];
dirs.forEach(dir => {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
});

// Static file serving
app.use('/uploads', express.static('uploads'));
app.use('/output', express.static('output'));
app.use('/downloads', express.static('downloads'));

// API Routes
app.use('/api', processRoutes);
app.use('/api/dashboard', dashboardRoutes);
app.use('/api/attendance', attendanceRoutes);
app.use('/api', authRoutes); // 👈 Login API

// Serve login page
app.get('/login', (req, res) => {
  res.sendFile(path.join(__dirname, '../frontend', 'login.html'));
});

app.get('/home', (req, res) => {
  res.sendFile(path.join(__dirname, '../frontend', 'index.html'));
});
// Serve dashboard page
app.get('/attendance',(req, res) => {
  res.sendFile(path.join(__dirname, '../frontend', 'attendance.html'));
});

app.get('/dashboard', (req, res) => {
  res.sendFile(path.join(__dirname, '../frontend', 'dashboard.html'));
});
// Default route (optional redirect to login)
app.get('/', (req, res) => {
  res.redirect('/login');
});

// Serve frontend static files
app.use(express.static(path.join(__dirname, '../frontend')));

// Error handling
app.use((error, req, res, next) => {
  console.error('Server Error:', error);
  res.status(500).json({
    error: 'Internal server error',
    message: error.message
  });
});

// 404 handler
app.use((req, res) => {
  res.status(404).json({ error: 'Route not found' });
});

app.listen(PORT, () => {
  console.log(`🚀 Server running on http://localhost:${PORT}`);
  console.log(`📁 Upload directory: ${path.resolve('uploads')}`);
  console.log(`📄 Output directory: ${path.resolve('output')}`);
  console.log(`🛡️ Login available at: http://localhost:${PORT}/login`);
  console.log(`🎯 Dashboard available at: http://localhost:${PORT}/dashboard`);
});

module.exports = app;
