/**
 * IFINITY Business Card Approval Workflow
 *
 * Employee -> Line Manager -> HR + Finance -> Marketing
 *
 * Install ONE spreadsheet "On form submit" trigger
 * for the handleFormSubmit function.
 */

const CONFIG = {
  APPROVER_EMAIL_FIELD: 'Email Address',
  REQUEST_SHEET: 'Form Responses 1',
  APPROVAL_SHEET: 'Form Responses 2',

  // ID from the approval Google Form URL
  APPROVAL_FORM_ID: '1jIhlQFNYZ2K97dxewI3dOjw9QtXD-A6W3lrVUeGWn38',

  // Replace with actual authorized addresses
  HR_EMAIL: 'qkrn239@gmail.com',
  FINANCE_EMAIL: 'vimukthi.nelanga@gmail.com',
  MARKETING_EMAIL: 'anuradha.bandaragoda@gmail.com',

  // Replace examples with actual manager names and emails.
  // Keys must match the request form dropdown exactly.
  MANAGERS: {
    'manager1': 'nelangabandaragoda@gmail.com',
    'manager2': 'manager2@ifinityglobal.com'
  },

  // Exact request form question titles

  EMPLOYEE_NAME_FIELD: 'Employee Name',
  EMAIL_FIELD: 'Employee Official Email',
  MANAGER_FIELD: "Line Manager's Name",
  REQUEST_TYPE_FIELD: 'Request Type',
  APPROVER_EMAIL_FIELD: 'Email Address',

};

const TRACKING_HEADERS = [
  'Request ID',
  'Approval Status',
  'Approved By',
  'Approval Date',
  'HR Status',
  'Finance Status',
  'Manager Comments',
  'HR Comments',
  'Finance Comments',
  'Printing Status',
  'Completion Date'
];

function setupWorkflow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const requestSheet =
    ss.getSheetByName(CONFIG.REQUEST_SHEET);

  const approvalSheet =
    ss.getSheetByName(CONFIG.APPROVAL_SHEET);

  if (!requestSheet || !approvalSheet) {
    throw new Error(
      'Check the request and approval response tab names.'
    );
  }

  const headers = getHeaders(requestSheet);

  TRACKING_HEADERS.forEach(header => {
    if (!headers.includes(header)) {
      requestSheet.getRange(
        1, requestSheet.getLastColumn() + 1
      ).setValue(header);
      headers.push(header);
    }
  });

  const form = FormApp.openById(
    CONFIG.APPROVAL_FORM_ID
  );

  ['Request ID', 'Decision', 'Comments']
    .forEach(title => {
      if (!form.getItems().some(
        item => item.getTitle().trim() === title
      )) {
        throw new Error(
          'Approval form is missing: ' + title
        );
      }
    });

  // Prevent duplicate workflow triggers.
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (
      trigger.getHandlerFunction() ===
      'handleFormSubmit'
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('handleFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  Logger.log('IFINITY workflow installed.');
}

function handleFormSubmit(e) {
  if (!e || !e.range) {
    throw new Error(
      'This function must run from a form submit trigger.'
    );
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = e.range.getSheet();
    const row = e.range.getRow();

    if (sheet.getName() === CONFIG.REQUEST_SHEET) {
      processNewRequest(sheet, row);
    } else if (
      sheet.getName() === CONFIG.APPROVAL_SHEET
    ) {
      processApproval(sheet, row);
    }
  } finally {
    lock.releaseLock();
  }
}

function processNewRequest(sheet, row) {
  const data = getRowData(sheet, row);

  // Prevent duplicate processing.
  if (data['Request ID']) return;


  const employeeName = String(
    data[CONFIG.EMPLOYEE_NAME_FIELD] || ''
  ).trim();

  const employeeEmail = normalizeEmail(
    data[CONFIG.EMAIL_FIELD]
  );

  const managerName = String(
    data[CONFIG.MANAGER_FIELD] || ''
  ).trim();

  const managerEmail = normalizeEmail(
    CONFIG.MANAGERS[managerName]
  );


  const requestId = 'BC-' +
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyyMMdd'
    ) + '-' + String(row).padStart(4, '0');

  setValue(sheet, row, 'Request ID', requestId);
  setValue(
    sheet, row, 'Printing Status', 'Not Started'
  );

  if (
    !employeeName ||
    !employeeEmail ||
    !managerEmail
  ) {
    setValue(
      sheet, row, 'Approval Status',
      'Configuration Error'
    );

    MailApp.sendEmail({
      to: CONFIG.MARKETING_EMAIL,
      subject: 'Action Required: ' + requestId,
      body:
        'A request has missing employee or ' +
        'manager information.\n\n' +
        'Check row ' + row +
        ' in the request spreadsheet.'
    });

    return;
  }

  setValue(
    sheet, row, 'Approval Status',
    'Pending Line Manager Approval'
  );

  setValue(
    sheet, row, 'HR Status', 'Not Started'
  );

  setValue(
    sheet, row, 'Finance Status', 'Not Started'
  );

  sendApprovalEmail(
    managerEmail,
    requestId,
    'LINE_MANAGER',
    getRowData(sheet, row)
  );

  sendStatusNotifications(
    requestId,
    getRowData(sheet, row),
    'SUBMITTED'
  );
}

function processApproval(sheet, row) {
  const response = getRowData(sheet, row);

  const requestId = String(
    response['Request ID'] || ''
  ).trim();

  const decision = String(
    response['Decision'] || ''
  ).trim();

  const comments = String(
    response['Comments'] || ''
  ).trim();

  const approverEmail = normalizeEmail(
    response[CONFIG.APPROVER_EMAIL_FIELD]
  );

  if (
    !requestId ||
    !['Approve', 'Reject'].includes(decision) ||
    !approverEmail
  ) {
    return;
  }

  const requestSheet =
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(CONFIG.REQUEST_SHEET);

  const requestRow = findRequestRow(
    requestSheet, requestId
  );

  if (!requestRow) {
    notifyInvalidApproval(
      approverEmail, requestId,
      'Request ID not found.'
    );
    return;
  }

  const data = getRowData(
    requestSheet, requestRow
  );

  const managerEmail = normalizeEmail(
    CONFIG.MANAGERS[
    data[CONFIG.MANAGER_FIELD]
    ]
  );

  const currentStatus =
    data['Approval Status'];

  const employeeEmail = normalizeEmail(
    data[CONFIG.EMAIL_FIELD]
  );

  const approved =
    decision === 'Approve';

  // Stage 1: Line Manager
  if (
    approverEmail === managerEmail &&
    currentStatus ===
    'Pending Line Manager Approval'
  ) {
    setValue(
      requestSheet, requestRow,
      'Approved By',
      approved ? approverEmail : ''
    );

    setValue(
      requestSheet, requestRow,
      'Approval Date',
      approved ? new Date() : ''
    );

    setValue(
      requestSheet, requestRow,
      'Manager Comments', comments
    );

    if (!approved) {
      rejectRequest(
        requestSheet, requestRow,
        requestId, employeeEmail,
        'Line Manager', comments
      );
      return;
    }

    setValue(
      requestSheet, requestRow,
      'Approval Status',
      'Pending HR and Finance Approval'
    );

    setValue(
      requestSheet, requestRow,
      'HR Status', 'Pending'
    );

    setValue(
      requestSheet, requestRow,
      'Finance Status', 'Pending'
    );

    const updated = getRowData(
      requestSheet, requestRow
    );

    sendApprovalEmail(
      CONFIG.HR_EMAIL,
      requestId, 'HR', updated
    );

    sendApprovalEmail(
      CONFIG.FINANCE_EMAIL,
      requestId, 'FINANCE', updated
    );

    // Notify Marketing and employee
    sendStatusNotifications(
      requestId,
      getRowData(requestSheet, requestRow),
      'MANAGER_APPROVED'
    );

    return;
  }

  // Stage 2: HR
  if (
    approverEmail ===
    normalizeEmail(CONFIG.HR_EMAIL) &&
    currentStatus ===
    'Pending HR and Finance Approval' &&
    data['HR Status'] === 'Pending'
  ) {
    setValue(
      requestSheet, requestRow,
      'HR Status',
      approved ? 'Approved' : 'Rejected'
    );

    setValue(
      requestSheet, requestRow,
      'HR Comments', comments
    );

    if (!approved) {
      rejectRequest(
        requestSheet, requestRow,
        requestId, employeeEmail,
        'HR', comments
      );
      return;
    }

    sendStatusNotifications(
      requestId,
      getRowData(requestSheet, requestRow),
      'HR_APPROVED'
    );

    checkFinalApproval(
      requestSheet, requestRow
    );
    return;
  }

  // Stage 2: Finance
  if (
    approverEmail ===
    normalizeEmail(CONFIG.FINANCE_EMAIL) &&
    currentStatus ===
    'Pending HR and Finance Approval' &&
    data['Finance Status'] === 'Pending'
  ) {
    setValue(
      requestSheet, requestRow,
      'Finance Status',
      approved ? 'Approved' : 'Rejected'
    );

    setValue(
      requestSheet, requestRow,
      'Finance Comments', comments
    );

    if (!approved) {
      rejectRequest(
        requestSheet, requestRow,
        requestId, employeeEmail,
        'Finance', comments
      );
      return;
    }

    sendStatusNotifications(
      requestId,
      getRowData(requestSheet, requestRow),
      'FINANCE_APPROVED'
    );

    checkFinalApproval(
      requestSheet, requestRow
    );
    return;
  }

  // Invalid, unauthorized, or duplicate decision.
  notifyInvalidApproval(
    approverEmail,
    requestId,
    'You are not the authorized pending approver, ' +
    'or this stage has already been completed.'
  );
}

function checkFinalApproval(sheet, row) {
  const data = getRowData(sheet, row);

  if (
    data['HR Status'] !== 'Approved' ||
    data['Finance Status'] !== 'Approved' ||
    data['Approval Status'] !==
    'Pending HR and Finance Approval'
  ) {
    return;
  }

  const requestId = data['Request ID'];

  setValue(
    sheet, row,
    'Approval Status', 'Fully Approved'
  );

  setValue(
    sheet, row,
    'Printing Status', 'Ready for Printing'
  );

  const details = formatRequestDetails(
    getRowData(sheet, row)
  );


  MailApp.sendEmail({
    to: CONFIG.MARKETING_EMAIL,
    subject:
      'APPROVED - Proceed with Printing | ' +
      requestId,
    body:
      'All approvals have been received.\n\n' +
      details +
      '\n\nYou may proceed with printing.',
    htmlBody: buildDepartmentEmailHtml(
      requestId,
      getRowData(sheet, row),
      '',
      'MARKETING'
    ),
    name: 'IFINITY | Business Card Requests'
  });


  const employeeEmail = normalizeEmail(
    data[CONFIG.EMAIL_FIELD]
  );

  if (employeeEmail) {
    MailApp.sendEmail({
      to: employeeEmail,
      subject:
        'Business Card Request Approved | ' +
        requestId,
      body:
        'Your business card request has received ' +
        'all required approvals and has been ' +
        'forwarded to Marketing for printing.\n\n' +
        'Request ID: ' + requestId
    });
  }
}

function rejectRequest(
  sheet, row, requestId,
  employeeEmail, rejectedBy, comments
) {
  setValue(
    sheet, row,
    'Approval Status',
    'Rejected by ' + rejectedBy
  );

  const message =
    'Business card request: ' + requestId +
    '\nRejected by: ' + rejectedBy +
    '\nComments: ' + (comments || 'None');

  MailApp.sendEmail({
    to: CONFIG.MARKETING_EMAIL,
    subject:
      'Business Card Request Rejected | ' +
      requestId,
    body: message
  });

  if (employeeEmail) {
    MailApp.sendEmail({
      to: employeeEmail,
      subject:
        'Business Card Request Rejected | ' +
        requestId,
      body:
        message +
        '\n\nPlease contact the relevant ' +
        'department for clarification.'
    });
  }
}



function sendApprovalEmail(
  recipient, requestId, stage, data
) {
  const link = buildApprovalLink(requestId);

  const stageName = {
    LINE_MANAGER: 'Line Manager',
    HR: 'HR',
    FINANCE: 'Finance'
  }[stage];

  const subject =
    'Approval Required | Business Card | ' +
    requestId;

  const body =
    'Dear ' + stageName + ',\n\n' +
    'A business card request requires your approval.\n\n' +
    formatRequestDetails(data) +
    '\n\nReview and submit your decision:\n' +
    link + '\n\n' +
    'IFINITY Business Card Workflow';

  const html = stage === 'LINE_MANAGER'
    ? buildManagerEmailHtml(
      requestId, data, link
    )
    : buildDepartmentEmailHtml(
      requestId, data, link, stage
    );

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    body: body,
    htmlBody: html,
    name: 'IFINITY | Business Card Requests'
  });
}




function buildApprovalLink(requestId) {
  const form = FormApp.openById(
    CONFIG.APPROVAL_FORM_ID
  );

  // Find Request ID regardless of extra spaces
  const requestItem = form.getItems().find(item =>
    item.getTitle().trim() === 'Request ID'
  );

  if (!requestItem) {
    throw new Error(
      'Approval form Request ID item missing. ' +
      'Available items: ' +
      form.getItems()
        .map(item => item.getTitle())
        .join(', ')
    );
  }

  // Ensure Request ID is a short-answer question
  if (
    requestItem.getType() !==
    FormApp.ItemType.TEXT
  ) {
    throw new Error(
      'Request ID must be a Short answer question.'
    );
  }

  // Generate prefilled approval form URL
  const response = form.createResponse();

  response.withItemResponse(
    requestItem.asTextItem()
      .createResponse(requestId)
  );

  return response.toPrefilledUrl();
}


function formatRequestDetails(data) {
  const excluded = new Set([
    'Timestamp',
    'Request ID',
    ...TRACKING_HEADERS
  ]);

  const lines = [
    'Request ID: ' + data['Request ID']
  ];

  Object.keys(data).forEach(key => {
    if (
      !excluded.has(key) &&
      data[key] !== '' &&
      data[key] !== null
    ) {
      lines.push(key + ': ' + data[key]);
    }
  });

  return lines.join('\n');
}

function findRequestRow(sheet, requestId) {
  const headers = getHeaders(sheet);
  const column =
    headers.indexOf('Request ID') + 1;

  if (!column) return null;

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return null;

  const values = sheet.getRange(
    2, column, lastRow - 1, 1
  ).getDisplayValues();

  const index = values.findIndex(
    item => item[0] === requestId
  );

  return index < 0 ? null : index + 2;
}

function getHeaders(sheet) {
  return sheet.getRange(
    1, 1, 1, sheet.getLastColumn()
  ).getDisplayValues()[0]
    .map(value => String(value).trim());
}


function getRowData(sheet, row) {
  const headers = getHeaders(sheet);

  const values = sheet.getRange(
    row, 1, 1, headers.length
  ).getValues()[0];

  const data = {};

  headers.forEach((header, index) => {
    data[header] = values[index];
  });

  return data;
}

function setValue(sheet, row, header, value) {
  const headers = getHeaders(sheet);

  const column =
    headers.indexOf(header) + 1;

  if (!column) {
    throw new Error(
      'Missing tracking column: ' + header
    );
  }

  sheet.getRange(row, column)
    .setValue(value);
}

function normalizeEmail(value) {
  return String(value || '')
    .trim()
    .toLowerCase();
}

function notifyInvalidApproval(
  email, requestId, reason
) {
  if (!email) return;

  MailApp.sendEmail({
    to: email,
    subject:
      'Approval Not Accepted | ' + requestId,
    body:
      'Your approval submission was not ' +
      'accepted.\n\nReason: ' + reason +
      '\n\nPlease contact Marketing ' +
      'if you believe this is an error.'
  });
}

function checkApprovalFormQuestions() {
  const form = FormApp.openById(
    CONFIG.APPROVAL_FORM_ID
  );

  Logger.log('Form: ' + form.getTitle());

  form.getItems().forEach(item => {
    Logger.log(
      'Question: [' + item.getTitle() + ']'
    );
  });
}




function debugRequestRow() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.REQUEST_SHEET);

  const row = sheet.getLastRow();
  const data = getRowData(sheet, row);

  Logger.log(
    'Employee Name: ' +
    JSON.stringify(data[CONFIG.EMPLOYEE_NAME_FIELD])
  );

  Logger.log(
    'Employee Email: ' +
    JSON.stringify(data[CONFIG.EMAIL_FIELD])
  );

  Logger.log(
    'Manager Name: ' +
    JSON.stringify(data[CONFIG.MANAGER_FIELD])
  );

  Logger.log(
    'Manager Email: ' +
    JSON.stringify(
      CONFIG.MANAGERS[
      String(data[CONFIG.MANAGER_FIELD] || '').trim()
      ]
    )
  );


  Logger.log('=== REQUEST DETAILS ===');

  Logger.log(
    'Request Type: ' +
    JSON.stringify(data['Request Type'])
  );

  Logger.log(
    'Card Name: ' +
    JSON.stringify(
      data['Name as it should appear on the card']
    )
  );

  Logger.log(
    'Card Email: ' +
    JSON.stringify(data['Email address'])
  );

  Logger.log(
    'Respondent Email: ' +
    JSON.stringify(data['Email Address'])
  );

  Logger.log(
    'Manager Name: ' +
    JSON.stringify(data["Line Manager's Name"])
  );

  Logger.log(
    'Manager Email Entered: ' +
    JSON.stringify(
      data["Line Manager's Email Address"]
    )
  );

  Logger.log(
    'Manager Email from Approved List: ' +
    JSON.stringify(
      CONFIG.MANAGERS[
      String(
        data["Line Manager's Name"] || ''
      ).trim()
      ]
    )
  );

  Logger.log('=== END ===');
}



function testApprovalLink() {
  const testId = 'BC-TEST-001';

  const link = buildApprovalLink(testId);

  Logger.log('Approval link generated:');
  Logger.log(link);
}


function escapeEmailHtml(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function buildManagerEmailHtml(requestId, data, approvalLink) {
  const e = escapeEmailHtml;

  const fields = [
    ['Employee Name', data['Employee Name']],
    ['Employee Email', data['Employee Official Email']],
    ['Request Type', data['Request Type']],
    ['Name on Card', data['Name as it should appear on the card']],
    ['Designation', data['Role / Designation']],
    ['Contact Number 1', data['Contact Number 1']],
    ['Contact Number 2', data['Contact Number 2']],
    ['Card Email', data['Email address']],
    ['Company Address', data['Company Address']],
    ['Quantity', data['How many units?']],
    ['Changes Required',
      data['Are there any changes required to your existing business card?']],
    ['Requested Updates', data['Mention what should be updated']],
    ['Priority', data['Priority']],
    ['Required Date', data['Due date']]
  ];

  const rows = fields
    .filter(([, value]) =>
      value !== '' && value !== null &&
      value !== undefined
    )
    .map(([label, value], index) => {
      const background = index % 2 === 0
        ? '#F8F8FB' : '#FFFFFF';

      const formatted = value instanceof Date
        ? Utilities.formatDate(
          value,
          Session.getScriptTimeZone(),
          'dd MMM yyyy'
        )
        : value;

      return `
        <tr style="background:${background};">
          <td style="padding:11px 14px;
                     color:#777C8B;
                     font-size:13px;
                     width:40%;
                     border-bottom:1px solid #ECECF2;">
            ${e(label)}
          </td>
          <td style="padding:11px 14px;
                     color:#20255B;
                     font-size:13px;
                     font-weight:600;
                     border-bottom:1px solid #ECECF2;">
            ${e(formatted)}
          </td>
        </tr>`;
    }).join('');

  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
</head>
<body style="margin:0;padding:24px 12px;
             background:#F1F2F6;
             font-family:Arial,Helvetica,sans-serif;">

  <table role="presentation" cellpadding="0"
         cellspacing="0" border="0"
         style="max-width:600px;width:100%;
                margin:0 auto;background:#FFFFFF;
                border-collapse:collapse;">

    <tr>
      <td style="background:#3A3A80;padding:24px;">
        <table role="presentation" width="100%">
          <tr>
            <td style="color:#FFFFFF;
                       font-size:23px;
                       font-weight:bold;">
              IFINITY
            </td>
            <td align="right"
                style="color:#FFFFFF;
                       font-size:10px;
                       letter-spacing:1px;">
              BUSINESS CARD REQUEST
            </td>
          </tr>
        </table>
      </td>
    </tr>

    <tr>
      <td style="padding:30px 26px;">

        <h1 style="font-size:24px;
                   color:#20255B;
                   margin:0 0 18px;">
          Approval Required
        </h1>

        <p style="color:#606579;
                  font-size:14px;
                  line-height:1.7;">
          Dear Line Manager,
        </p>

        <p style="color:#606579;
                  font-size:14px;
                  line-height:1.7;">
          A business card request has been submitted
          and requires your approval before proceeding
          to HR and Finance.
        </p>

        <table role="presentation" width="100%"
               style="margin:24px 0;">
          <tr>
            <td>
              <div style="font-size:11px;
                          color:#777C8B;">
                REQUEST REFERENCE
              </div>
              <div style="font-size:16px;
                          font-weight:bold;
                          color:#20255B;
                          margin-top:5px;">
                ${e(requestId)}
              </div>
            </td>
            <td align="right">
              <span style="background:#FFF0DE;
                           color:#A05A17;
                           padding:7px 10px;
                           font-size:11px;
                           font-weight:bold;">
                PENDING APPROVAL
              </span>
            </td>
          </tr>
        </table>

        <h2 style="font-size:16px;
                   color:#20255B;
                   margin:24px 0 14px;">
          Employee &amp; Request Details
        </h2>

        <table role="presentation" width="100%"
               cellpadding="0" cellspacing="0"
               style="border-collapse:collapse;
                      border:1px solid #ECECF2;">
          ${rows}
        </table>

        <div style="background:#F5F5FA;
                    padding:16px;
                    margin:24px 0;
                    color:#606579;
                    font-size:13px;
                    line-height:1.6;">
          <strong style="color:#3A3A80;">
            Next Step
          </strong>
          <p style="margin:7px 0 0;">
            Please review the request and submit
            your approval decision using the
            link below.
          </p>
        </div>

        <table role="presentation" align="center"
               cellpadding="0" cellspacing="0"
               style="margin:25px auto;">
          <tr>
            <td bgcolor="#3A3A80"
                style="padding:15px 28px;">
              <a href="${e(approvalLink)}"
                 style="color:#FFFFFF;
                        text-decoration:none;
                        font-size:14px;
                        font-weight:bold;">
                Review Request
              </a>
            </td>
          </tr>
        </table>

        <p style="color:#85899A;
                  text-align:center;
                  font-size:12px;
                  line-height:1.6;">
          Upon approval, this request will be
          forwarded to HR and Finance.
        </p>
      </td>
    </tr>

    <tr>
      <td align="center"
          style="background:#F8F8FB;
                 padding:22px;">
        <div style="font-size:13px;
                    font-weight:bold;
                    color:#3A3A80;">
          IFINITY Marketing Division
        </div>
        <div style="font-size:11px;
                    color:#888D9C;
                    margin-top:6px;">
          Automated Business Card Approval Workflow
        </div>
      </td>
    </tr>
  </table>
</body>
</html>`;
}



function buildDepartmentEmailHtml(
  requestId, data, approvalLink, department
) {
  const e = escapeEmailHtml;

  const rows = [
    ['Employee Name', data['Employee Name']],
    ['Employee Email', data['Employee Official Email']],
    ['Request Type', data['Request Type']],
    ['Name on Card', data['Name as it should appear on the card']],
    ['Designation', data['Role / Designation']],
    ['Contact Number 1', data['Contact Number 1']],
    ['Contact Number 2', data['Contact Number 2']],
    ['Card Email', data['Email address']],
    ['Company Address', data['Company Address']],
    ['Quantity', data['How many units?']],
    ['Changes Required',
      data['Are there any changes required to your existing business card?']],
    ['Requested Updates', data['Mention what should be updated']],
    ['Priority', data['Priority']],
    ['Required Date', data['Due date']]
  ]
    .filter(([, value]) =>
      value !== '' && value !== null &&
      value !== undefined
    )
    .map(([label, value], index) => {
      const bg = index % 2 === 0
        ? '#F8F8FB' : '#FFFFFF';

      const formatted = value instanceof Date
        ? Utilities.formatDate(
          value,
          Session.getScriptTimeZone(),
          'dd MMM yyyy'
        )
        : value;

      return `
      <tr style="background:${bg};">
        <td style="padding:11px 14px;
                   color:#777C8B;
                   font-size:13px;
                   width:40%;
                   border-bottom:1px solid #ECECF2;">
          ${e(label)}
        </td>
        <td style="padding:11px 14px;
                   color:#20255B;
                   font-size:13px;
                   font-weight:600;
                   border-bottom:1px solid #ECECF2;">
          ${e(formatted)}
        </td>
      </tr>`;
    }).join('');

  const isMarketing = department === 'MARKETING';

  const title = isMarketing
    ? 'Approved - Proceed with Printing'
    : 'Approval Required';

  const intro = isMarketing
    ? 'The following business card request has received all required approvals and is ready for printing.'
    : 'The following business card request has been approved by the Line Manager and requires your review.';

  const status = isMarketing
    ? 'FULLY APPROVED'
    : 'PENDING ' + department + ' APPROVAL';

  const hrStatus = isMarketing
    ? 'Approved'
    : String(data['HR Status'] || 'Pending');

  const financeStatus = isMarketing
    ? 'Approved'
    : String(data['Finance Status'] || 'Pending');

  const statusColor = isMarketing
    ? '#23734A' : '#A05A17';

  const statusBg = isMarketing
    ? '#E7F5EC' : '#FFF0DE';

  const action = isMarketing
    ? `
      <div style="background:#E7F5EC;
                  color:#23734A;
                  padding:18px;
                  margin-top:24px;
                  font-size:14px;
                  line-height:1.6;">
        <strong>Authorized for Printing</strong>
        <p style="margin:7px 0 0;">
          All mandatory approvals have been received.
          Please proceed with printing arrangements.
        </p>
      </div>`
    : `
      <table role="presentation"
             align="center"
             cellpadding="0"
             cellspacing="0"
             style="margin:26px auto;">
        <tr>
          <td bgcolor="#3A3A80"
              style="padding:15px 28px;">
            <a href="${e(approvalLink)}"
               style="color:#FFFFFF;
                      text-decoration:none;
                      font-size:14px;
                      font-weight:bold;">
              Review Request
            </a>
          </td>
        </tr>
      </table>
      <p style="color:#85899A;
                text-align:center;
                font-size:12px;">
        Marketing will be notified after both
        HR and Finance approvals are completed.
      </p>`;

  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
</head>
<body style="margin:0;
             padding:24px 12px;
             background:#F1F2F6;
             font-family:Arial,Helvetica,sans-serif;">

<table role="presentation"
       cellpadding="0"
       cellspacing="0"
       border="0"
       style="max-width:600px;
              width:100%;
              margin:0 auto;
              background:#FFFFFF;
              border-collapse:collapse;">

  <tr>
    <td style="background:#3A3A80;
               padding:24px;">
      <table role="presentation" width="100%">
        <tr>
          <td style="color:#FFFFFF;
                     font-size:23px;
                     font-weight:bold;">
            IFINITY
          </td>
          <td align="right"
              style="color:#FFFFFF;
                     font-size:10px;
                     letter-spacing:1px;">
            BUSINESS CARD REQUEST
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <tr>
    <td style="padding:30px 26px;">

      <h1 style="font-size:24px;
                 color:#20255B;
                 margin:0 0 18px;">
        ${title}
      </h1>

      <p style="color:#606579;
                font-size:14px;
                line-height:1.7;">
        Dear ${isMarketing ? 'Marketing Team' : e(department) + ' Team'},
      </p>

      <p style="color:#606579;
                font-size:14px;
                line-height:1.7;">
        ${intro}
      </p>

      <table role="presentation"
             width="100%"
             style="margin:24px 0;">
        <tr>
          <td>
            <div style="font-size:11px;
                        color:#777C8B;">
              REQUEST REFERENCE
            </div>
            <div style="font-size:16px;
                        font-weight:bold;
                        color:#20255B;
                        margin-top:5px;">
              ${e(requestId)}
            </div>
          </td>
          <td align="right">
            <span style="background:${statusBg};
                         color:${statusColor};
                         padding:7px 10px;
                         font-size:11px;
                         font-weight:bold;">
              ${e(status)}
            </span>
          </td>
        </tr>
      </table>

      <h2 style="font-size:16px;
                 color:#20255B;">
        Employee &amp; Request Details
      </h2>

      <table role="presentation"
             width="100%"
             cellpadding="0"
             cellspacing="0"
             style="border-collapse:collapse;
                    border:1px solid #ECECF2;">
        ${rows}
      </table>

      <div style="background:#F5F5FA;
                  padding:18px;
                  margin-top:24px;
                  font-size:13px;
                  color:#606579;">

        <strong style="color:#20255B;
                       font-size:14px;">
          Approval Summary
        </strong>

        <table role="presentation"
               width="100%"
               style="margin-top:12px;
                      font-size:13px;
                      color:#606579;">
          <tr>
            <td style="padding:6px 0;">Line Manager</td>
            <td align="right"
                style="color:#23734A;">
              Approved
            </td>
          </tr>
          <tr>
            <td style="padding:6px 0;">HR</td>
            <td align="right">${e(hrStatus)}</td>
          </tr>
          <tr>
            <td style="padding:6px 0;">Finance</td>
            <td align="right">${e(financeStatus)}</td>
          </tr>
        </table>
      </div>

      ${action}

    </td>
  </tr>

  <tr>
    <td align="center"
        style="background:#F8F8FB;
               padding:22px;">
      <div style="font-size:13px;
                  font-weight:bold;
                  color:#3A3A80;">
        IFINITY Marketing Division
      </div>
      <div style="font-size:11px;
                  color:#888D9C;
                  margin-top:6px;">
        Automated Business Card Approval Workflow
      </div>
    </td>
  </tr>
</table>
</body>
</html>`;
}


function testDepartmentEmailTemplates() {
  const testEmail =
    Session.getEffectiveUser().getEmail();

  const requestId = 'BC-EMAIL-PREVIEW';

  const sampleData = {
    'Employee Name': 'Sample Employee',
    'Employee Official Email': 'sample@example.com',
    'Request Type': 'New Business Card Request',
    'Name as it should appear on the card':
      'Sample Employee',
    'Role / Designation': 'Marketing Executive',
    'Contact Number 1': '+94 XX XXX XXXX',
    'Email address': 'sample@example.com',
    'Company Address': 'IFINITY Office',
    'How many units?': '100',
    'HR Status': 'Pending',
    'Finance Status': 'Pending'
  };

  ['HR', 'FINANCE', 'MARKETING'].forEach(stage => {
    const html = buildDepartmentEmailHtml(
      requestId,
      sampleData,
      'https://forms.google.com',
      stage
    );

    MailApp.sendEmail({
      to: testEmail,
      subject:
        'TEST | IFINITY ' + stage +
        ' Email Template',
      body: 'HTML email preview.',
      htmlBody: html
    });
  });

  Logger.log('Three test emails submitted.');
}


function buildStatusEmailHtml(requestId, data, event, audience) {
  const e = escapeEmailHtml;

  const events = {
    SUBMITTED: {
      title: 'Business Card Request Received',
      status: 'Pending Manager Approval',
      message: 'The request has been submitted successfully and is awaiting Line Manager approval.',
      next: 'The Line Manager will review the request.'
    },
    MANAGER_APPROVED: {
      title: 'Line Manager Approval Completed',
      status: 'Manager Approved',
      message: 'The Line Manager has approved the business card request.',
      next: 'HR and Finance will now review the request.'
    },
    HR_APPROVED: {
      title: 'HR Approval Completed',
      status: 'HR Approved',
      message: 'HR has approved the business card request.',
      next: 'The request will proceed once all remaining approvals are completed.'
    },
    FINANCE_APPROVED: {
      title: 'Finance Approval Completed',
      status: 'Finance Approved',
      message: 'Finance has approved the business card request.',
      next: 'The request will proceed once all remaining approvals are completed.'
    },
    FULLY_APPROVED: {
      title: 'Business Card Request Fully Approved',
      status: 'Fully Approved',
      message: 'All required approvals have been completed.',
      next: 'Marketing will coordinate the printing process.'
    },
    REJECTED: {
      title: 'Business Card Request Rejected',
      status: 'Rejected',
      message: 'The request was rejected during the approval process.',
      next: 'Please review the decision and contact the relevant department if clarification is required.'
    }
  };

  const info = events[event];
  if (!info) throw new Error('Unknown notification event: ' + event);

  const isRejected = event === 'REJECTED';
  const isComplete = event === 'FULLY_APPROVED';

  const statusBg = isRejected ? '#FCE8E8'
    : isComplete ? '#E7F5EC' : '#FFF0DE';

  const statusColor = isRejected ? '#A12B2B'
    : isComplete ? '#23734A' : '#A05A17';

  const fields = [
    ['Employee Name', data['Employee Name']],
    ['Request Type', data['Request Type']],
    ['Quantity', data['How many units?']],
    ['Current Status', data['Approval Status']],
    ['Line Manager',
      data["Line Manager's Name"]]
  ];

  const rows = fields
    .filter(([, value]) =>
      value !== '' && value !== null &&
      value !== undefined
    )
    .map(([label, value], index) => `
      <tr style="background:${index % 2 ? '#FFFFFF' : '#F8F8FB'};">
        <td style="padding:11px 14px;color:#777C8B;
                   font-size:13px;width:42%;
                   border-bottom:1px solid #ECECF2;">
          ${e(label)}
        </td>
        <td style="padding:11px 14px;color:#20255B;
                   font-size:13px;font-weight:600;
                   border-bottom:1px solid #ECECF2;">
          ${e(value)}
        </td>
      </tr>
    `).join('');

  const managerStatus = data['Approved By']
    ? 'Approved'
    : String(data['Approval Status'] || '').startsWith('Rejected by Line Manager')
      ? 'Rejected' : 'Pending';

  const hrStatus = String(data['HR Status'] || 'Not Started');
  const financeStatus = String(data['Finance Status'] || 'Not Started');

  const progress = [
    ['Line Manager', managerStatus],
    ['HR', hrStatus],
    ['Finance', financeStatus]
  ].map(([label, status]) => `
    <tr>
      <td style="padding:7px 0;color:#606579;font-size:13px;">
        ${e(label)}
      </td>
      <td align="right"
          style="padding:7px 0;font-size:13px;
                 font-weight:bold;
                 color:${status === 'Approved' ? '#23734A'
      : status === 'Rejected' ? '#A12B2B'
        : '#A05A17'};">
        ${e(status)}
      </td>
    </tr>
  `).join('');

  const greeting = audience === 'MARKETING'
    ? 'Dear Marketing Team,'
    : 'Dear Employee,';

  return `
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:24px 12px;background:#F1F2F6;
             font-family:Arial,Helvetica,sans-serif;">

<table role="presentation" cellpadding="0" cellspacing="0"
       width="100%" style="max-width:600px;margin:0 auto;
                          background:#FFFFFF;border-collapse:collapse;">

  <tr>
    <td style="background:#3A3A80;padding:24px;color:#FFFFFF;">
      <table role="presentation" width="100%">
        <tr>
          <td style="font-size:23px;font-weight:bold;color:#FFFFFF;">
            IFINITY
          </td>
          <td align="right" style="font-size:10px;color:#FFFFFF;
                                   letter-spacing:1px;">
            BUSINESS CARD REQUEST
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <tr>
    <td style="padding:30px 26px;">
      <h1 style="font-size:23px;color:#20255B;margin:0 0 18px;">
        ${e(info.title)}
      </h1>

      <p style="color:#606579;font-size:14px;line-height:1.7;">
        ${greeting}
      </p>

      <p style="color:#606579;font-size:14px;line-height:1.7;">
        ${e(info.message)}
      </p>

      <table role="presentation" width="100%"
             style="margin:24px 0;">
        <tr>
          <td>
            <div style="font-size:11px;color:#777C8B;">
              REQUEST REFERENCE
            </div>
            <div style="font-size:16px;font-weight:bold;
                        color:#20255B;margin-top:5px;">
              ${e(requestId)}
            </div>
          </td>
          <td align="right">
            <span style="background:${statusBg};
                         color:${statusColor};
                         padding:7px 10px;
                         font-size:11px;font-weight:bold;">
              ${e(info.status)}
            </span>
          </td>
        </tr>
      </table>

      <h2 style="font-size:16px;color:#20255B;">
        Request Summary
      </h2>

      <table role="presentation" width="100%"
             cellpadding="0" cellspacing="0"
             style="border-collapse:collapse;
                    border:1px solid #ECECF2;">
        ${rows}
      </table>

      <div style="background:#F5F5FA;padding:18px;margin-top:24px;">
        <strong style="color:#20255B;font-size:14px;">
          Approval Progress
        </strong>
        <table role="presentation" width="100%"
               style="margin-top:12px;">
          ${progress}
        </table>
      </div>

      <div style="padding:16px 0;color:#606579;
                  font-size:13px;line-height:1.7;">
        <strong style="color:#3A3A80;">Next Step</strong>
        <p style="margin:6px 0;">${e(info.next)}</p>
      </div>

    </td>
  </tr>

  <tr>
    <td align="center" style="background:#F8F8FB;padding:22px;">
      <div style="color:#3A3A80;font-size:13px;font-weight:bold;">
        IFINITY Marketing Division
      </div>
      <div style="color:#888D9C;font-size:11px;margin-top:6px;">
        Automated Business Card Approval Workflow
      </div>
    </td>
  </tr>
</table>
</body>
</html>`;
}

function sendStatusNotifications(requestId, data, event) {
  const employeeEmail = normalizeEmail(
    data[CONFIG.EMAIL_FIELD]
  );

  const subjectMap = {
    SUBMITTED: 'Request Received',
    MANAGER_APPROVED: 'Line Manager Approved',
    HR_APPROVED: 'HR Approved',
    FINANCE_APPROVED: 'Finance Approved',
    FULLY_APPROVED: 'Request Fully Approved',
    REJECTED: 'Request Rejected'
  };

  if (!subjectMap[event]) {
    throw new Error('Unknown status event: ' + event);
  }

  const subject =
    subjectMap[event] + ' | Business Card | ' + requestId;

  const recipients = [
    ...(event === 'FULLY_APPROVED' ? [] : [{
      email: CONFIG.MARKETING_EMAIL,
      audience: 'MARKETING'
    }]),
    {
      email: employeeEmail,
      audience: 'EMPLOYEE'
    }
  ];

  recipients.forEach(recipient => {
    if (!recipient.email) return;

    MailApp.sendEmail({
      to: recipient.email,
      subject: subject,
      body:
        'Business card request ' + requestId +
        '\nStatus: ' + subjectMap[event],
      htmlBody: buildStatusEmailHtml(
        requestId,
        data,
        event,
        recipient.audience
      ),
      name: 'IFINITY | Business Card Requests'
    });
  });
}
