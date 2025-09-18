const express = require('express');
const router = express.Router();
const jwt = require('jsonwebtoken');
const { jwtSecret } = require('../config');

const users=[{
  id:1,
  username:'admin',
  password:'admin123',
  name:'admin1',
  department:'it',
  role:'admin'
}]

// POST /api/login
router.post('/login', async (req, res) => {
  const {username, password } = req.body;

  try {
   const user=users.find(u=>u.username===username);

   if(!user){
    return res.status(401).json({success:false,message:'invalid credentials'});
   }

   if(password!==user.password){
    return res.status(401).json({success:false,message:'invalid credentials'});
   }

    // ✅ Generate JWT Token
    const token = jwt.sign(
      { id: user.id, username: user.username, role: 'user' }, // payload
      jwtSecret,
      { expiresIn: '2h' } // token expires in 2 hours
    );

    res.json({
      success: true,
      message: 'Login successful',
      token,
      user: {
        id: user.id,
        name: user.name,
        department: user.department,
        username: user.username
      }
    });

  } catch (err) {
    console.error('Login error:', err);
    res.status(500).json({ success: false, message: 'Server error' });
  }
});


module.exports = router;
