// app_auth.gs
// User identity and access lookup helpers.

function getUserEmail() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (e) {
    return '';
  }
}

function getUserPermission() {
  try {
    const email = (Session.getActiveUser().getEmail() || '').toString().trim().toLowerCase();
    const ss = getSpreadsheet();
    let usersSheet = resolveSheet(CONFIG.SHEETS.USERS);
    if (!usersSheet) {
      usersSheet = ss.insertSheet(CONFIG.SHEETS.USERS);
      usersSheet.appendRow(['Email', 'สำนัก/กอง', 'Role']);
      return null;
    }
    const data = usersSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if ((data[i][0] || '').toString().trim().toLowerCase() === email) {
        return {
          email: (data[i][0] || '').toString().trim(),
          department: (data[i][1] || '').toString().trim(),
          role: normalizeRoleValue(data[i][2]),
          rawRole: (data[i][2] || '').toString().trim()
        };
      }
    }
    return null;
  } catch (error) {
    handleError('getUserPermission', error);
    return null;
  }
}
