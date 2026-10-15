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

function rows_(name) {
  setup_();
  const sheet = SpreadsheetApp.getActive().getSheetByName(name);
  return sheet.getLastRow() < 2
    ? []
    : sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
}

function student_(barcode) {
  const wanted = normalize_(barcode);
  const rows = rows_(S.STUDENTS);
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
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

function snack_(barcode) {
  const wanted = normalize_(barcode);
  const rows = rows_(S.SNACKS);
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
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

function balance_(barcode) {
  const student = student_(barcode);
  if (!student) return null;

  let reloads = 0;
  let purchases = 0;
  let adjustments = 0;
  const wanted = normalize_(barcode);

  rows_(S.RELOADS).forEach(row => {
    if (normalize_(row[1]) === wanted) reloads += Number(row[3]) || 0;
  });
  rows_(S.TX).forEach(row => {
    if (normalize_(row[1]) === wanted) purchases += Number(row[5]) || 0;
  });
  rows_(S.ADJUSTMENTS).forEach(row => {
    if (normalize_(row[1]) === wanted) adjustments += Number(row[3]) || 0;
  });

  return student.opening + reloads + adjustments - purchases;
}

function remaining_(barcode) {
  const snack = snack_(barcode);
  if (!snack || snack.qty === null) return null;
  let sold = 0;
  const wanted = normalize_(barcode);
  rows_(S.TX).forEach(row => {
    if (normalize_(row[3]) === wanted) sold++;
  });
  return snack.qty - sold;
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
  const student = student_(barcode);
  if (!student) return {ok: false, message: 'Student not found.'};
  if (!student.active) return {ok: false, message: 'Student account is inactive.'};
  return {
    ok: true,
    student: {
      barcode: student.barcode,
      name: student.name,
      grade: student.grade,
      balance: balance_(barcode)
    }
  };
}

function addStudent(data) {
  setup_();
  const barcode = normalize_(data.barcode);
  const name = normalize_(data.name);
  const grade = normalize_(data.grade);
  const opening = nonnegativeWholeNumber_(data.opening === '' ? 0 : data.opening, 'Starting points');

  if (!barcode) throw Error('Scan the student ID first.');
  if (!name) throw Error('Enter the student name.');
  if (student_(barcode)) throw Error('That barcode is already assigned. Reactivate the existing account if needed.');

  SpreadsheetApp.getActive().getSheetByName(S.STUDENTS)
    .appendRow([barcode, name, grade, opening, 'Yes', new Date()]);
  return {message: name + ' added with ' + opening + ' points.'};
}

function setStudentActive(data) {
  setup_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);
  try {
    const barcode = normalize_(data.barcode);
    const active = data.active === true || String(data.active).toLowerCase() === 'true';
    const reason = normalize_(data.reason) || (active ? 'Account reactivated' : 'Account deactivated');
    const staff = normalize_(data.staff);
    const student = student_(barcode);

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
    const barcode = normalize_(data.barcode);
    const points = positiveWholeNumber_(data.points, 'Points');
    const paid = data.paid === '' ? 0 : Number(data.paid);
    const staff = normalize_(data.staff);
    const student = student_(barcode);

    if (!student) throw Error('Student not found.');
    if (!student.active) throw Error('Student account is inactive. Reactivate it before adding points.');
    if (!Number.isFinite(paid) || paid < 0) throw Error('Amount paid cannot be negative.');

    SpreadsheetApp.getActive().getSheetByName(S.RELOADS)
      .appendRow([new Date(), barcode, student.name, points, paid, staff]);
    return {message: 'Added ' + points + ' points to ' + student.name + '.', balance: balance_(barcode)};
  } finally {
    lock.releaseLock();
  }
}

function subtractPoints(data) {
  setup_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);
  try {
    const barcode = normalize_(data.barcode);
    const points = positiveWholeNumber_(data.points, 'Points to subtract');
    const reason = normalize_(data.reason) || 'Manual point correction';
    const staff = normalize_(data.staff);
    const student = student_(barcode);

    if (!student) throw Error('Student not found.');
    if (!student.active) throw Error('Student account is inactive. Reactivate it before changing points.');

    const before = balance_(barcode);
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
  const barcode = normalize_(data.barcode);
  const name = normalize_(data.name);
  const points = positiveWholeNumber_(data.points, 'Point cost');
  const cost = data.cost === '' ? '' : Number(data.cost);
  const qty = data.qty === '' ? '' : nonnegativeWholeNumber_(data.qty, 'Starting quantity');

  if (!barcode) throw Error('Scan the snack barcode first.');
  if (!name) throw Error('Enter the snack name.');
  if (cost !== '' && (!Number.isFinite(cost) || cost < 0)) throw Error('Unit cost cannot be negative.');
  if (snack_(barcode)) throw Error('That snack barcode is already registered.');

  SpreadsheetApp.getActive().getSheetByName(S.SNACKS)
    .appendRow([barcode, name, points, cost, qty, 'Yes', new Date()]);
  return {message: name + ' registered for ' + points + ' point' + (points === 1 ? '' : 's') + '.'};
}

function checkout(data) {
  setup_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);
  try {
    const studentBarcode = normalize_(data.studentBarcode);
    const snackBarcode = normalize_(data.snackBarcode);
    const staff = normalize_(data.staff);
    const student = student_(studentBarcode);
    const snack = snack_(snackBarcode);

    if (!student) throw Error('Student not found.');
    if (!student.active) throw Error('Student account is inactive.');
    if (!snack) throw Error('Snack not found.');
    if (!snack.active) throw Error('Snack is inactive.');

    const before = balance_(studentBarcode);
    const remaining = remaining_(snackBarcode);
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
  const students = rows_(S.STUDENTS).filter(row => normalize_(row[0]));
  const snacks = rows_(S.SNACKS).filter(row => normalize_(row[0]));
  const transactions = rows_(S.TX);
  let outstanding = 0;
  students.forEach(row => { outstanding += balance_(row[0]) || 0; });
  const activeStudents = students.filter(row => String(row[4]).toLowerCase() !== 'no').length;
  return {students: activeStudents, snacks: snacks.length, sales: transactions.length, points: outstanding};
}

function getStudents() {
  return rows_(S.STUDENTS)
    .filter(row => normalize_(row[0]))
    .map(row => ({
      barcode: row[0],
      name: row[1],
      grade: row[2],
      balance: balance_(row[0]),
      active: String(row[4]).toLowerCase() !== 'no'
    }));
}

function getSnacks() {
  return rows_(S.SNACKS)
    .filter(row => normalize_(row[0]))
    .map(row => ({barcode: row[0], name: row[1], points: Number(row[2]) || 0, remaining: remaining_(row[0])}));
}

function getSales() {
  return rows_(S.TX).slice(-100).reverse()
    .map(row => ({time: row[0], student: row[2], snack: row[4], points: row[5], balance: row[7], staff: row[8]}));
}

function getAdjustments() {
  return rows_(S.ADJUSTMENTS).slice(-100).reverse()
    .map(row => ({time: row[0], student: row[2], change: row[3], balance: row[5], reason: row[6], staff: row[7]}));
}
