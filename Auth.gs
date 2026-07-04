/* =========================
   AUTH MODULE
========================= */

function checkLogin(username, password) {
  if (
    String(username).trim() === ERP_CONFIG.LOGIN_USERNAME &&
    String(password).trim() === ERP_CONFIG.LOGIN_PASSWORD
  ) {
    return {
      success: true,
      name: 'Store Admin'
    };
  }

  return {
    success: false,
    message: 'Invalid username or password'
  };
}
