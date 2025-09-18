document.getElementById('loginForm').addEventListener('submit', async (e) => {
  e.preventDefault();

  const username = document.getElementById('username').value.trim();
  const password = document.getElementById('password').value.trim();

  const response = await fetch('/api/login', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ username, password })
  });

  const result = await response.json();
if (result.success) {
  // Store JWT in localStorage for later use
  localStorage.setItem('token', result.token);

  // Redirect to AMC homepage
  window.location.href = '/home';
 

}

else {
    document.getElementById('error').textContent = result.message || 'Invalid credentials';
  }
});
