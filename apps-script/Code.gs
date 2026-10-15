const S = {
  STUDENTS: 'Students',
  SNACKS: 'Snacks',
  RELOADS: 'Reloads',
  TX: 'Transactions',
  ADJUSTMENTS: 'Point Adjustments',
  STATUS_LOG: 'Student Status Log'
};

function doGet() {
  setup_();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('FBLA Snack Pass')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setup_() {
  const ss = SpreadsheetApp.getActive();
  makeSheet_(ss, S.STUDENTS, ['Student Barcode', 'Student Name', 'Grade', 'Opening Points', 'Active', 'Created At']);
  makeSheet_(ss, S.SNACKS, ['Snack Barcode', 'Snack Name', 'Point Cost', 'Unit Cost', 'Starting Quantity', 'Active', 'Created At']);
  makeSheet_(ss, S.RELOADS, ['Timestamp', 'Student Barcode', 'Student Name', 'Points Added', 'Amount Paid', 'Staff']);
  makeSheet_(ss, S.TX, ['Timestamp', 'Student Barcode', 'Student Name', 'Snack Barcode', 'Snack Name', 'Points Deducted', 'Balance Before', 'Balance After', 'Staff']);
  makeSheet_(ss, S.ADJUSTMENTS, ['Timestamp', 'Student Barcode', 'Student Name', 'Points Change', 'Balance Before', 'Balance After', 'Reason', 'Staff']);
  makeSheet_(ss, S.STATUS_LOG, ['Timestamp', 'Student Barcode', 'Student Name', 'Status', 'Reason', 'Staff']);
}

function makeSheet_(ss, name, headers) {
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (!sheet.getLastRow()) {
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#17365D')
      .setFontColor('#fff');
    sheet.setFrozenRows(1);
  }
}

function normalize_(value) {
  return String(value == null ? '' : value).trim();
}

function sheetRows_(sheet) {
  const lastRow = sheet.getLastRow();
  return lastRow < 2
    ? []
    : sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getDisplayValues();
}

function rows_(name) {
  setup_();
  return sheetRows_(SpreadsheetApp.getActive().getSheetByName(name));
}

// Load all data needed by a request once. Reusing this snapshot avoids the
// repeated full-sheet reads that made scans slow as history grew.
function snapshot_() {
  setup_();
  const ss = SpreadsheetApp.getActive();
  return {
    students: sheetRows_(ss.getSheetByName(S.STUDENTS)),
    snacks: sheetRows_(ss.getSheetByName(S.SNACKS)),
    reloads: sheetRows_(ss.getSheetByName(S.RELOADS)),
    transactions: sheetRows_(ss.getSheetByName(S.TX)),
    adjustments: sheetRows_(ss.getSheetByName(S.ADJUSTMENTS))
  };
}

function studentFrom_(data, barcode) {
  const wanted = normalize_(barcode);
  for (let i = 0; i < data.students.length; i++) {
    const row = data.students[i];
    if (normalize_(row[0]) === wanted) {
      return {
        row: i + 2,
        barcode: row[0],
        name: row[1],
        grade: row[2],
        opening: Number(row[3]) || 0,
        active: String(row[4]).toLowerCase() !== 'no'
      };
    }
  }
  return null;
}

function snackFrom_(data, barcode) {
  const wanted = normalize_(barcode);
  for (let i = 0; i < data.snacks.length; i++) {
    const row = data.snacks[i];
    if (normalize_(row[0]) === wanted) {
      return {
        barcode: row[0],
        name: row[1],
        points: Number(row[2]) || 0,
        cost: row[3] === '' ? null : Number(row[3]),
        qty: row[4] === '' ? null : Number(row[4]),
        active: String(row[5]).toLowerCase() !== 'no'
      };
    }
  }
  return null;
}

function balanceMap_(data) {
  const balances = Object.create(null);

  data.students.forEach(row => {
    const barcode = normalize_(row[0]);
    if (barcode) balances[barcode] = Number(row[3]) || 0;
  });
  data.reloads.forEach(row => {
    const barcode = normalize_(row[1]);
    if (barcode in balances) balances[barcode] += Number(row[3]) || 0;
  });
  data.adjustments.forEach(row => {
    const barcode = normalize_(row[1]);
    if (barcode in balances) balances[barcode] += Number(row[3]) || 0;
  });
  data.transactions.forEach(row => {
    const barcode = normalize_(row[1]);
    if (barcode in balances) balances[barcode] -= Number(row[5]) || 0;
  });

  return balances;
}

function inventoryMap_(data) {
  const remaining = Object.create(null);

  data.snacks.forEach(row => {
    const barcode = normalize_(row[0]);
    if (barcode) remaining[barcode] = row[4] === '' ? null : Number(row[4]);
  });
  data.transactions.forEach(row => {
    const barcode = normalize_(row[3]);
    if (barcode in remaining && remaining[barcode] !== null) remaining[barcode]--;
  });

  return remaining;
}

function positiveWholeNumber_(value, label) {
  const number = Number(value);
  if (!Number.isFinite(number) || number <= 0 || Math.floor(number) !== number) {
    throw Error(label + ' must be a whole number greater than zero.');
  }
  return number;
}

function nonnegativeWholeNumber_(value, label) {
  const number = Number(value);
  if (!Number.isFinite(number) || number < 0 || Math.floor(number) !== number) {
    throw Error(label + ' must be a whole number of zero or more.');
  }
  return number;
}

function lookupStudent(barcode) {
  const data = snapshot_();
  const student = studentFrom_(data, barcode);
  if (!student) return {ok: false, message: 'Student not found.'};
  if (!student.active) return {ok: false, message: 'Student account is inactive.'};
  const balances = balanceMap_(data);
  return {
    ok: true,
    student: {
      barcode: student.barcode,
      name: student.name,
      grade: student.grade,
      balance: balances[normalize_(student.barcode)]
    }
  };
}

function addStudent(data) {
  setup_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);
  try {
    const snapshot = snapshot_();
    const barcode = normalize_(data.barcode);
    const name = normalize_(data.name);
    const grade = normalize_(data.grade);
    const opening = nonnegativeWholeNumber_(data.opening === '' ? 0 : data.opening, 'Starting points');

    if (!barcode) throw Error('Scan the student ID first.');
    if (!name) throw Error('Enter the student name.');
    if (studentFrom_(snapshot, barcode)) throw Error('That barcode is already assigned. Reactivate the existing account if needed.');

    SpreadsheetApp.getActive().getSheetByName(S.STUDENTS)
      .appendRow([barcode, name, grade, opening, 'Yes', new Date()]);
    return {message: name + ' added with ' + opening + ' points.'};
  } finally {
    lock.releaseLock();
  }
}

function setStudentActive(data) {
  setup_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);
  try {
    const snapshot = snapshot_();
    const barcode = normalize_(data.barcode);
    const active = data.active === true || String(data.active).toLowerCase() === 'true';
    const reason = normalize_(data.reason) || (active ? 'Account reactivated' : 'Account deactivated');
    const staff = normalize_(data.staff);
    const student = studentFrom_(snapshot, barcode);

    if (!student) throw Error('Student not found.');
    if (student.active === active) {
      throw Error(student.name + ' is already ' + (active ? 'active.' : 'inactive.'));
    }

    SpreadsheetApp.getActive().getSheetByName(S.STUDENTS)
      .getRange(student.row, 5)
      .setValue(active ? 'Yes' : 'No');
    SpreadsheetApp.getActive().getSheetByName(S.STATUS_LOG)
      .appendRow([new Date(), student.barcode, student.name, active ? 'Active' : 'Inactive', reason, staff]);

    return {
      message: student.name + ' was ' + (active ? 'reactivated.' : 'deactivated. Their history was kept.'),
      active: active
    };
  } finally {
    lock.releaseLock();
  }
}

function loadPoints(data) {
  setup_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);
  try {
    const snapshot = snapshot_();
    const barcode = normalize_(data.barcode);
    const points = positiveWholeNumber_(data.points, 'Points');
    const paid = data.paid === '' ? 0 : Number(data.paid);
    const staff = normalize_(data.staff);
    const student = studentFrom_(snapshot, barcode);

    if (!student) throw Error('Student not found.');
    if (!student.active) throw Error('Student account is inactive. Reactivate it before adding points.');
    if (!Number.isFinite(paid) || paid < 0) throw Error('Amount paid cannot be negative.');

    const before = balanceMap_(snapshot)[normalize_(barcode)];
    SpreadsheetApp.getActive().getSheetByName(S.RELOADS)
      .appendRow([new Date(), barcode, student.name, points, paid, staff]);
    return {message: 'Added ' + points + ' points to ' + student.name + '.', balance: before + points};
  } finally {
    lock.releaseLock();
  }
}

function subtractPoints(data) {
  setup_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);
  try {
    const snapshot = snapshot_();
    const barcode = normalize_(data.barcode);
    const points = positiveWholeNumber_(data.points, 'Points to subtract');
    const reason = normalize_(data.reason) || 'Manual point correction';
    const staff = normalize_(data.staff);
    const student = studentFrom_(snapshot, barcode);

    if (!student) throw Error('Student not found.');
    if (!student.active) throw Error('Student account is inactive. Reactivate it before changing points.');

    const before = balanceMap_(snapshot)[normalize_(barcode)];
    if (before < 0) throw Error(student.name + ' already has an invalid negative balance. Correct the sheet before continuing.');
    if (points > before) {
      throw Error('Cannot subtract ' + points + ' points. ' + student.name + ' only has ' + before + '.');
    }

    const after = before - points;
    SpreadsheetApp.getActive().getSheetByName(S.ADJUSTMENTS)
      .appendRow([new Date(), barcode, student.name, -points, before, after, reason, staff]);
    return {message: 'Subtracted ' + points + ' points from ' + student.name + '.', balance: after};
  } finally {
    lock.releaseLock();
  }
}

function addSnack(data) {
  setup_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);
  try {
    const snapshot = snapshot_();
    const barcode = normalize_(data.barcode);
    const name = normalize_(data.name);
    const points = positiveWholeNumber_(data.points, 'Point cost');
    const cost = data.cost === '' ? '' : Number(data.cost);
    const qty = data.qty === '' ? '' : nonnegativeWholeNumber_(data.qty, 'Starting quantity');

    if (!barcode) throw Error('Scan the snack barcode first.');
    if (!name) throw Error('Enter the snack name.');
    if (cost !== '' && (!Number.isFinite(cost) || cost < 0)) throw Error('Unit cost cannot be negative.');
    if (snackFrom_(snapshot, barcode)) throw Error('That snack barcode is already registered.');

    SpreadsheetApp.getActive().getSheetByName(S.SNACKS)
      .appendRow([barcode, name, points, cost, qty, 'Yes', new Date()]);
    return {message: name + ' registered for ' + points + ' point' + (points === 1 ? '' : 's') + '.'};
  } finally {
    lock.releaseLock();
  }
}

function checkout(data) {
  setup_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);
  try {
    const snapshot = snapshot_();
    const studentBarcode = normalize_(data.studentBarcode);
    const snackBarcode = normalize_(data.snackBarcode);
    const staff = normalize_(data.staff);
    const student = studentFrom_(snapshot, studentBarcode);
    const snack = snackFrom_(snapshot, snackBarcode);

    if (!student) throw Error('Student not found.');
    if (!student.active) throw Error('Student account is inactive.');
    if (!snack) throw Error('Snack not found.');
    if (!snack.active) throw Error('Snack is inactive.');

    const before = balanceMap_(snapshot)[normalize_(studentBarcode)];
    const remaining = inventoryMap_(snapshot)[normalize_(snackBarcode)];
    if (remaining !== null && remaining <= 0) throw Error(snack.name + ' is out of stock.');
    if (before < snack.points) throw Error(student.name + ' only has ' + before + ' points.');

    const after = before - snack.points;
    SpreadsheetApp.getActive().getSheetByName(S.TX)
      .appendRow([new Date(), studentBarcode, student.name, snackBarcode, snack.name, snack.points, before, after, staff]);
    return {snack: snack.name, points: snack.points, balance: after};
  } finally {
    lock.releaseLock();
  }
}

function dashboard() {
  const data = snapshot_();
  const balances = balanceMap_(data);
  let outstanding = 0;
  data.students.forEach(row => {
    const barcode = normalize_(row[0]);
    if (barcode) outstanding += balances[barcode] || 0;
  });
  const activeStudents = data.students.filter(row => normalize_(row[0]) && String(row[4]).toLowerCase() !== 'no').length;
  const snacks = data.snacks.filter(row => normalize_(row[0])).length;
  return {students: activeStudents, snacks: snacks, sales: data.transactions.length, points: outstanding};
}

function getStudents() {
  const data = snapshot_();
  const balances = balanceMap_(data);
  return data.students
    .filter(row => normalize_(row[0]))
    .map(row => ({
      barcode: row[0],
      name: row[1],
      grade: row[2],
      balance: balances[normalize_(row[0])],
      active: String(row[4]).toLowerCase() !== 'no'
    }));
}

function getSnacks() {
  const data = snapshot_();
  const remaining = inventoryMap_(data);
  return data.snacks
    .filter(row => normalize_(row[0]))
    .map(row => ({
      barcode: row[0],
      name: row[1],
      points: Number(row[2]) || 0,
      remaining: remaining[normalize_(row[0])]
    }));
}

function getSales() {
  return rows_(S.TX).slice(-100).reverse()
    .map(row => ({time: row[0], student: row[2], snack: row[4], points: row[5], balance: row[7], staff: row[8]}));
}

function getAdjustments() {
  return rows_(S.ADJUSTMENTS).slice(-100).reverse()
    .map(row => ({time: row[0], student: row[2], change: row[3], balance: row[5], reason: row[6], staff: row[7]}));
}
