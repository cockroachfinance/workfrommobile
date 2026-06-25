function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const action = e.parameter.action;
    const redirect = e.parameter.redirect;
    const format = e.parameter.format; // ← ADD THIS LINE
    
    // Check if headers exist
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Timestamp',
        'Name',
        'Mobile',
        'Email',
        'Password',
        'Status',
        'URL'
      ]);
    }
    
    if (action === 'login') {
      const mobile = e.parameter.mobile;
      const password = e.parameter.password;
      
      const data = sheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        
        if (row[2] == mobile && row[4] == password) {
          const status = row[5] || 'Pending';
          const redirectUrl = row[6] || 'https://example.com/dashboard';
          
          // --- PENDING STATUS ---
          if (status.toLowerCase() === 'pending') {
            if (format === 'json') {
              return ContentService
                .createTextOutput(JSON.stringify({
                  success: false,
                  message: 'Your account is pending approval. Please wait for admin to activate.',
                  status: 'pending'
                }))
                .setMimeType(ContentService.MimeType.JSON);
            }
            return HtmlService.createHtmlOutput(
              '<!DOCTYPE html><html><head><script>alert("Your account is pending approval. Please wait for admin to activate."); window.history.back();</script></head><body>Account pending</body></html>'
            );
          }
          
          // --- BLOCKED STATUS ---
          if (status.toLowerCase() === 'blocked' || status.toLowerCase() === 'inactive') {
            if (format === 'json') {
              return ContentService
                .createTextOutput(JSON.stringify({
                  success: false,
                  message: 'Your account has been blocked. Please contact support.',
                  status: 'blocked'
                }))
                .setMimeType(ContentService.MimeType.JSON);
            }
            return HtmlService.createHtmlOutput(
              '<!DOCTYPE html><html><head><script>alert("Your account has been blocked. Please contact support."); window.history.back();</script></head><body>Account blocked</body></html>'
            );
          }
          
          // --- SUCCESS (Approved/Active) ---
          if (format === 'json') {
            return ContentService
              .createTextOutput(JSON.stringify({
                success: true,
                message: 'Login successful',
                redirectUrl: redirectUrl,
                status: status
              }))
              .setMimeType(ContentService.MimeType.JSON);
          }
          
          return HtmlService.createHtmlOutput(
            '<!DOCTYPE html><html><head><script>window.location.href="' + redirectUrl + '";</script></head><body>Redirecting...</body></html>'
          );
        }
      }
      
      // --- LOGIN FAILED ---
      if (format === 'json') {
        return ContentService
          .createTextOutput(JSON.stringify({
            success: false,
            message: 'Invalid mobile number or password. Please try again.'
          }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      
      return HtmlService.createHtmlOutput(
        '<!DOCTYPE html><html><head><script>alert("Invalid mobile number or password. Please try again."); window.history.back();</script></head><body>Invalid credentials</body></html>'
      );
    }
    
    // --- UPDATE ACTION (for admin) ---
    if (action === 'update') {
      const mobile = e.parameter.mobile;
      const newStatus = e.parameter.status;
      const newUrl = e.parameter.url;
      
      const data = sheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[2] == mobile) {
          if (newStatus) {
            sheet.getRange(i + 1, 6).setValue(newStatus);
          }
          if (newUrl) {
            sheet.getRange(i + 1, 7).setValue(newUrl);
          }
          
          return ContentService
            .createTextOutput(JSON.stringify({
              success: true,
              message: 'User updated successfully'
            }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
      
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          message: 'User not found'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: 'Invalid action'
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: err.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
