/**
 * PROJECT INVENTORY MANAGEMENT SYSTEM
 * File: EmailService.gs
 *
 * Handles:
 * - Approval request email
 * - Approve and Reject links
 * - Secure token-based approval action
 * - Approval email resend
 * - Email HTML template
 * - Web app approval page
 *
 * ApprovalService.gs automatically calls:
 * sendApprovalRequestEmail_(data)
 *
 * IMPORTANT:
 * Your main doGet(e) must route approval links to:
 *
 * if (
 *   e &&
 *   e.parameter &&
 *   e.parameter.page === 'approval'
 * ) {
 *   return handleApprovalEmailPage(e);
 * }
 */

const EMAIL_SERVICE_CONFIG = Object.freeze({
  APP_NAME:
    'Project Inventory Management System',

  APPROVAL_PAGE:
    'approval',

  ACTION: Object.freeze({
    APPROVE: 'approve',
    REJECT: 'reject'
  }),

  SUBJECT_PREFIX:
    '[Inventory Approval]',

  SENDER_NAME:
    'Inventory Management System',

  SUPPORT_EMAIL:
    'mis@anushagroup.com',

  MAX_REMARK_LENGTH:
    1000
});


/**
 * Returns deployed web app URL.
 *
 * Priority:
 * 1. APP_CONFIG.WEB_APP_URL
 * 2. Script Property WEB_APP_URL
 * 3. Current deployment URL
 */
function getApprovalWebAppUrl_() {
  let url = '';

  if (
    typeof APP_CONFIG !== 'undefined' &&
    APP_CONFIG.WEB_APP_URL
  ) {
    url = normalizeText_(
      APP_CONFIG.WEB_APP_URL
    );
  }

  if (!url) {
    url = normalizeText_(
      PropertiesService
        .getScriptProperties()
        .getProperty('WEB_APP_URL')
    );
  }

  if (!url) {
    try {
      url = normalizeText_(
        ScriptApp.getService().getUrl()
      );
    } catch (error) {
      url = '';
    }
  }

  if (!url) {
    throw new Error(
      'Web App URL is not configured. Deploy the Apps Script as a Web App, then run setApprovalWebAppUrl().'
    );
  }

  return url;
}


/**
 * Saves deployed web app URL.
 *
 * Run manually after deployment:
 *
 * setApprovalWebAppUrl(
 *   'https://script.google.com/macros/s/DEPLOYMENT_ID/exec'
 * )
 */
function setApprovalWebAppUrl(webAppUrl) {
  const url = normalizeText_(webAppUrl);

 if (
  !url ||
  url.indexOf('/exec') === -1
) {
  throw new Error(
    'Enter a valid deployed Apps Script Web App URL ending with /exec.'
  );
}

  PropertiesService
    .getScriptProperties()
    .setProperty(
      'WEB_APP_URL',
      url
    );

  Logger.log(
    'Web App URL saved successfully: ' +
    url
  );

  return successResponse_(
    'Web App URL saved successfully.',
    {
      webAppUrl: url
    }
  );
}


/**
 * ApprovalService.gs hook.
 *
 * Expected input:
 * {
 *   requestId,
 *   token,
 *   tokenExpiry,
 *   approverEmail,
 *   sourceProject,
 *   destinationProject,
 *   referenceNo,
 *   requestedBy,
 *   requestedByEmail,
 *   remarks,
 *   items
 * }
 */
function sendApprovalRequestEmail_(data) {
  validateApprovalEmailData_(data);

  const webAppUrl =
    getApprovalWebAppUrl_();

  const approveUrl =
    buildApprovalActionUrl_(
      webAppUrl,
      data.token,
      EMAIL_SERVICE_CONFIG
        .ACTION.APPROVE
    );

  const rejectUrl =
    buildApprovalActionUrl_(
      webAppUrl,
      data.token,
      EMAIL_SERVICE_CONFIG
        .ACTION.REJECT
    );

  const subject =
    EMAIL_SERVICE_CONFIG.SUBJECT_PREFIX +
    ' ' +
    data.requestId +
    ' | ' +
    data.sourceProject.projectName +
    ' → ' +
    data.destinationProject.projectName;

  const htmlBody =
    buildApprovalEmailHtml_({
      requestId:
        data.requestId,

      tokenExpiry:
        data.tokenExpiry,

      approverEmail:
        data.approverEmail,

      sourceProject:
        data.sourceProject,

      destinationProject:
        data.destinationProject,

      referenceNo:
        data.referenceNo,

      requestedBy:
        data.requestedBy,

      requestedByEmail:
        data.requestedByEmail,

      remarks:
        data.remarks,

      items:
        data.items,

      approveUrl:
        approveUrl,

      rejectUrl:
        rejectUrl
    });

  const plainBody =
    buildApprovalEmailText_({
      requestId:
        data.requestId,

      tokenExpiry:
        data.tokenExpiry,

      sourceProject:
        data.sourceProject,

      destinationProject:
        data.destinationProject,

      referenceNo:
        data.referenceNo,

      requestedBy:
        data.requestedBy,

      requestedByEmail:
        data.requestedByEmail,

      remarks:
        data.remarks,

      items:
        data.items,

      approveUrl:
        approveUrl,

      rejectUrl:
        rejectUrl
    });

  MailApp.sendEmail({
  to: data.approverEmail,
  subject: subject,
  body: plainBody,
  htmlBody: htmlBody,
  name: EMAIL_SERVICE_CONFIG.SENDER_NAME
});

  try {
    addTransactionLog_(
      getCurrentUserEmail_(),
      'EMAIL',
      'SEND APPROVAL EMAIL',
      data.requestId,
      {
        approverEmail:
          data.approverEmail,
        tokenExpiry:
          formatDateTime_(
            data.tokenExpiry
          )
      }
    );
  } catch (logError) {
    console.error(
      'Approval email log failed:',
      logError
    );
  }

  return successResponse_(
    'Approval email sent successfully.',
    {
      requestId:
        data.requestId,

      approverEmail:
        data.approverEmail,

      tokenExpiry:
        formatDateTime_(
          data.tokenExpiry
        )
    }
  );
}


/**
 * Validates approval email payload.
 */
function validateApprovalEmailData_(data) {
  if (
    !data ||
    typeof data !== 'object'
  ) {
    throw new Error(
      'Approval email data is required.'
    );
  }

  const required = [
    'requestId',
    'token',
    'tokenExpiry',
    'approverEmail',
    'sourceProject',
    'destinationProject',
    'referenceNo',
    'requestedBy',
    'items'
  ];

  required.forEach(function (field) {
    if (
      !Object.prototype
        .hasOwnProperty.call(
          data,
          field
        ) ||
      data[field] === '' ||
      data[field] === null ||
      typeof data[field] ===
        'undefined'
    ) {
      throw new Error(
        'Approval email field is missing: ' +
        field
      );
    }
  });

  if (
    !isValidEmail_(
      data.approverEmail
    )
  ) {
    throw new Error(
      'Invalid approver email: ' +
      data.approverEmail
    );
  }

  if (
    !Array.isArray(data.items) ||
    data.items.length === 0
  ) {
    throw new Error(
      'Approval email requires at least one item.'
    );
  }

  return true;
}


/**
 * Creates approval/rejection URL.
 */
function buildApprovalActionUrl_(
  webAppUrl,
  token,
  action
) {
  return (
    webAppUrl +
    '?page=' +
    encodeURIComponent(
      EMAIL_SERVICE_CONFIG
        .APPROVAL_PAGE
    ) +
    '&action=' +
    encodeURIComponent(action) +
    '&token=' +
    encodeURIComponent(token)
  );
}


/**
 * Builds approval email HTML.
 */
function buildApprovalEmailHtml_(data) {
  const rows = data.items
    .map(function (item, index) {
      const quantity =
        toNumber_(
          item.quantity,
          0
        );

      const rate =
        toNumber_(
          item.rate,
          0
        );

      const amount =
        roundTwo_(
          quantity * rate
        );

      return (
        '<tr>' +
          '<td style="padding:8px;border:1px solid #d1d5db;text-align:center;">' +
            (index + 1) +
          '</td>' +
          '<td style="padding:8px;border:1px solid #d1d5db;">' +
            escapeHtml_(
              item.skuCode || ''
            ) +
          '</td>' +
          '<td style="padding:8px;border:1px solid #d1d5db;">' +
            escapeHtml_(
              item.skuName || ''
            ) +
          '</td>' +
          '<td style="padding:8px;border:1px solid #d1d5db;text-align:right;">' +
            escapeHtml_(quantity) +
          '</td>' +
          '<td style="padding:8px;border:1px solid #d1d5db;">' +
            escapeHtml_(
              item.sku &&
              item.sku.uom
                ? item.sku.uom
                : item.uom || ''
            ) +
          '</td>' +
          '<td style="padding:8px;border:1px solid #d1d5db;text-align:right;">' +
            escapeHtml_(
              formatCurrency_(rate)
            ) +
          '</td>' +
          '<td style="padding:8px;border:1px solid #d1d5db;text-align:right;">' +
            escapeHtml_(
              formatCurrency_(amount)
            ) +
          '</td>' +
        '</tr>'
      );
    })
    .join('');

  const totalQuantity =
    roundTwo_(
      data.items.reduce(
        function (total, item) {
          return total +
            toNumber_(
              item.quantity,
              0
            );
        },
        0
      )
    );

  const totalAmount =
    roundTwo_(
      data.items.reduce(
        function (total, item) {
          return total +
            (
              toNumber_(
                item.quantity,
                0
              ) *
              toNumber_(
                item.rate,
                0
              )
            );
        },
        0
      )
    );

  return (
    '<div style="font-family:Arial,sans-serif;background:#f3f4f6;padding:24px;color:#111827;">' +
      '<div style="max-width:760px;margin:0 auto;background:#ffffff;border:1px solid #e5e7eb;border-radius:12px;overflow:hidden;">' +

        '<div style="background:#1f4e79;color:#ffffff;padding:20px 24px;">' +
          '<div style="font-size:20px;font-weight:700;">' +
            escapeHtml_(
              EMAIL_SERVICE_CONFIG
                .APP_NAME
            ) +
          '</div>' +
          '<div style="margin-top:6px;font-size:14px;">Cross-project customized stock approval</div>' +
        '</div>' +

        '<div style="padding:24px;">' +
          '<p style="margin-top:0;">Hello,</p>' +
          '<p>A customized inventory outward request requires your approval.</p>' +

          '<table style="width:100%;border-collapse:collapse;margin:18px 0;">' +
            emailInfoRow_(
              'Request ID',
              data.requestId
            ) +
            emailInfoRow_(
              'Reference No',
              data.referenceNo
            ) +
            emailInfoRow_(
              'Source Project',
              data.sourceProject.projectName +
              ' (' +
              data.sourceProject.projectId +
              ')'
            ) +
            emailInfoRow_(
              'Destination Project',
              data.destinationProject.projectName +
              ' (' +
              data.destinationProject.projectId +
              ')'
            ) +
            emailInfoRow_(
              'Requested By',
              data.requestedBy
            ) +
            emailInfoRow_(
              'Requester Email',
              data.requestedByEmail || ''
            ) +
            emailInfoRow_(
              'Token Expiry',
              formatDateTime_(
                data.tokenExpiry
              )
            ) +
          '</table>' +

          '<table style="width:100%;border-collapse:collapse;margin-top:20px;">' +
            '<thead>' +
              '<tr style="background:#f9fafb;">' +
                '<th style="padding:8px;border:1px solid #d1d5db;">#</th>' +
                '<th style="padding:8px;border:1px solid #d1d5db;">SKU Code</th>' +
                '<th style="padding:8px;border:1px solid #d1d5db;">SKU Name</th>' +
                '<th style="padding:8px;border:1px solid #d1d5db;">Qty</th>' +
                '<th style="padding:8px;border:1px solid #d1d5db;">Unit</th>' +
                '<th style="padding:8px;border:1px solid #d1d5db;">Rate</th>' +
                '<th style="padding:8px;border:1px solid #d1d5db;">Amount</th>' +
              '</tr>' +
            '</thead>' +
            '<tbody>' +
              rows +
            '</tbody>' +
          '</table>' +

          '<div style="margin-top:16px;text-align:right;">' +
            '<div><strong>Total Quantity:</strong> ' +
              escapeHtml_(
                totalQuantity
              ) +
            '</div>' +
            '<div style="margin-top:6px;"><strong>Total Amount:</strong> ' +
              escapeHtml_(
                formatCurrency_(
                  totalAmount
                )
              ) +
            '</div>' +
          '</div>' +

          (
            data.remarks
              ? (
                  '<div style="margin-top:18px;padding:12px;background:#fff7ed;border-left:4px solid #f97316;">' +
                    '<strong>Remarks:</strong> ' +
                    escapeHtml_(
                      data.remarks
                    ) +
                  '</div>'
                )
              : ''
          ) +

          '<div style="margin-top:26px;text-align:center;">' +
            '<a href="' +
              escapeHtml_(
                data.approveUrl
              ) +
            '" style="display:inline-block;background:#15803d;color:#ffffff;text-decoration:none;padding:12px 22px;border-radius:7px;font-weight:700;margin:4px;">Approve</a>' +

            '<a href="' +
              escapeHtml_(
                data.rejectUrl
              ) +
            '" style="display:inline-block;background:#b91c1c;color:#ffffff;text-decoration:none;padding:12px 22px;border-radius:7px;font-weight:700;margin:4px;">Reject</a>' +
          '</div>' +

          '<p style="font-size:12px;color:#6b7280;margin-top:22px;">The links are secure and will expire on ' +
            escapeHtml_(
              formatDateTime_(
                data.tokenExpiry
              )
            ) +
            '.</p>' +
        '</div>' +

        '<div style="padding:14px 24px;background:#f9fafb;color:#6b7280;font-size:12px;text-align:center;">' +
          'This is an automated inventory approval email.' +
        '</div>' +

      '</div>' +
    '</div>'
  );
}


/**
 * Email information row.
 */
function emailInfoRow_(label, value) {
  return (
    '<tr>' +
      '<td style="padding:8px;border:1px solid #e5e7eb;background:#f9fafb;width:180px;font-weight:700;">' +
        escapeHtml_(label) +
      '</td>' +
      '<td style="padding:8px;border:1px solid #e5e7eb;">' +
        escapeHtml_(value) +
      '</td>' +
    '</tr>'
  );
}


/**
 * Builds plain-text approval email.
 */
function buildApprovalEmailText_(data) {
  const itemText = data.items
    .map(function (item, index) {
      return (
        (index + 1) +
        '. ' +
        (item.skuCode || '') +
        ' - ' +
        (item.skuName || '') +
        ' | Qty: ' +
        toNumber_(
          item.quantity,
          0
        ) +
        ' | Rate: ' +
        formatCurrency_(
          toNumber_(
            item.rate,
            0
          )
        )
      );
    })
    .join('\n');

  return [
    EMAIL_SERVICE_CONFIG.APP_NAME,
    '',
    'Cross-project customized stock approval is required.',
    '',
    'Request ID: ' +
      data.requestId,
    'Reference No: ' +
      data.referenceNo,
    'Source Project: ' +
      data.sourceProject.projectName,
    'Destination Project: ' +
      data.destinationProject.projectName,
    'Requested By: ' +
      data.requestedBy,
    'Requester Email: ' +
      (data.requestedByEmail || ''),
    'Token Expiry: ' +
      formatDateTime_(
        data.tokenExpiry
      ),
    '',
    'Items:',
    itemText,
    '',
    'Remarks: ' +
      (data.remarks || ''),
    '',
    'Approve:',
    data.approveUrl,
    '',
    'Reject:',
    data.rejectUrl
  ].join('\n');
}


/**
 * Resends approval email for a pending request.
 */
function resendApprovalRequestEmail(
  requestId
) {
  return safeExecute_(function () {
    requirePermission_('canApprove');

    const records =
      getApprovalRows_(
        requestId
      );

    if (records.length === 0) {
      throw new Error(
        'Approval request not found: ' +
        requestId
      );
    }

    assertSingleApprovalStatus_(
      records,
      APPROVAL_SERVICE_CONFIG
        .STATUS.PENDING
    );

    const first = records[0];

    const result =
      sendApprovalRequestEmail_({
        requestId:
          first.requestId,

        token:
          first.approvalToken,

        tokenExpiry:
          first.tokenExpiry,

        approverEmail:
          first.approverEmail,

        sourceProject: {
          projectId:
            first.sourceProjectId,
          projectName:
            first.sourceProject
        },

        destinationProject: {
          projectId:
            first.destinationProjectId,
          projectName:
            first.destinationProject
        },

        referenceNo:
          first.referenceNo,

        requestedBy:
          first.requestedBy,

        requestedByEmail:
          first.requestedByEmail,

        remarks:
          first.remarks,

        items:
          records.map(
            function (record) {
              return {
                skuId:
                  record.skuId,
                skuCode:
                  record.skuCode,
                skuName:
                  record.skuName,
                quantity:
                  record.quantity,
                uom:
                  record.unit,
                rate:
                  record.rate
              };
            }
          )
      });

    return result;
  }, 'Unable to resend approval email.');
}


/**
 * Handles email approval web page.
 *
 * Add this route to your main doGet(e):
 *
 * if (
 *   e.parameter.page === 'approval'
 * ) {
 *   return handleApprovalEmailPage(e);
 * }
 */
function handleApprovalEmailPage(e) {
  const action =
    normalizeLower_(
      e &&
      e.parameter
        ? e.parameter.action
        : ''
    );

  const token =
    normalizeText_(
      e &&
      e.parameter
        ? e.parameter.token
        : ''
    );

  if (
    !token ||
    [
      EMAIL_SERVICE_CONFIG
        .ACTION.APPROVE,

      EMAIL_SERVICE_CONFIG
        .ACTION.REJECT
    ].indexOf(action) === -1
  ) {
    return buildApprovalMessagePage_(
      'Invalid Link',
      'The approval link is invalid or incomplete.',
      false
    );
  }

  let request;

  try {
    const records =
      getApprovalRowsByToken_(
        token
      );

    validateApprovalTokenRecords_(
      records
    );

    request =
      mapApprovalRequest_(records);

  } catch (error) {
    return buildApprovalMessagePage_(
      'Approval Link Unavailable',
      error.message ||
      'This approval link is no longer available.',
      false
    );
  }

  const html =
    buildApprovalActionPageHtml_({
      action:
        action,

      token:
        token,

      request:
        request
    });

  return HtmlService
    .createHtmlOutput(html)
    .setTitle(
      action ===
      EMAIL_SERVICE_CONFIG
        .ACTION.APPROVE
        ? 'Approve Inventory Request'
        : 'Reject Inventory Request'
    )
    .setXFrameOptionsMode(
      HtmlService
        .XFrameOptionsMode
        .ALLOWALL
    );
}


/**
 * HTML approval/rejection confirmation page.
 */
function buildApprovalActionPageHtml_(
  data
) {
  const isApprove =
    data.action ===
    EMAIL_SERVICE_CONFIG
      .ACTION.APPROVE;

  const title = isApprove
    ? 'Approve Inventory Request'
    : 'Reject Inventory Request';

  const buttonText = isApprove
    ? 'Confirm Approval'
    : 'Confirm Rejection';

  const buttonColor = isApprove
    ? '#15803d'
    : '#b91c1c';

  const remarksLabel = isApprove
    ? 'Approval Remarks (Optional)'
    : 'Rejection Reason (Required)';

  return (
    '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
      '<base target="_top">' +
      '<meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<style>' +
        'body{font-family:Arial,sans-serif;background:#f3f4f6;margin:0;padding:20px;color:#111827;}' +
        '.card{max-width:650px;margin:30px auto;background:#fff;border-radius:12px;padding:24px;box-shadow:0 8px 30px rgba(0,0,0,.08);}' +
        'h1{font-size:22px;margin:0 0 18px;color:#1f4e79;}' +
        '.info{background:#f9fafb;border:1px solid #e5e7eb;border-radius:8px;padding:14px;margin-bottom:18px;}' +
        '.row{margin:7px 0;}' +
        'label{display:block;font-weight:700;margin:15px 0 7px;}' +
        'textarea{width:100%;min-height:110px;box-sizing:border-box;padding:11px;border:1px solid #cbd5e1;border-radius:7px;font-family:inherit;}' +
        'button{width:100%;border:0;border-radius:7px;padding:13px;color:#fff;font-size:15px;font-weight:700;cursor:pointer;background:' +
          buttonColor +
        ';}' +
        'button:disabled{opacity:.6;cursor:not-allowed;}' +
        '#message{margin-top:15px;padding:11px;border-radius:7px;display:none;}' +
        '.success{display:block!important;background:#dcfce7;color:#166534;}' +
        '.error{display:block!important;background:#fee2e2;color:#991b1b;}' +
      '</style>' +
    '</head>' +
    '<body>' +
      '<div class="card">' +
        '<h1>' +
          escapeHtml_(title) +
        '</h1>' +

        '<div class="info">' +
          '<div class="row"><strong>Request ID:</strong> ' +
            escapeHtml_(
              data.request.requestId
            ) +
          '</div>' +
          '<div class="row"><strong>Source:</strong> ' +
            escapeHtml_(
              data.request.sourceProject
            ) +
          '</div>' +
          '<div class="row"><strong>Destination:</strong> ' +
            escapeHtml_(
              data.request.destinationProject
            ) +
          '</div>' +
          '<div class="row"><strong>Total Quantity:</strong> ' +
            escapeHtml_(
              data.request.totalQuantity
            ) +
          '</div>' +
        '</div>' +

        '<label for="remarks">' +
          escapeHtml_(
            remarksLabel
          ) +
        '</label>' +

        '<textarea id="remarks" maxlength="' +
          EMAIL_SERVICE_CONFIG
            .MAX_REMARK_LENGTH +
        '"></textarea>' +

        '<button id="submitButton" onclick="submitAction()">' +
          escapeHtml_(
            buttonText
          ) +
        '</button>' +

        '<div id="message"></div>' +
      '</div>' +

      '<script>' +
        'function submitAction(){' +
          'var button=document.getElementById("submitButton");' +
          'var message=document.getElementById("message");' +
          'var remarks=document.getElementById("remarks").value.trim();' +

          (
            isApprove
              ? ''
              : (
                  'if(!remarks){' +
                    'message.className="error";' +
                    'message.textContent="Rejection reason is required.";' +
                    'return;' +
                  '}'
                )
          ) +

          'button.disabled=true;' +
          'button.textContent="Processing...";' +
          'message.className="";' +
          'message.style.display="none";' +

          'google.script.run' +
            '.withSuccessHandler(function(result){' +
              'if(result&&result.success){' +
                'message.className="success";' +
                'message.textContent=result.message;' +
                'button.textContent="Completed";' +
              '}else{' +
                'message.className="error";' +
                'message.textContent=result&&result.message?result.message:"Operation failed.";' +
                'button.disabled=false;' +
                'button.textContent=' +
                  JSON.stringify(buttonText) +
                ';' +
              '}' +
            '})' +
            '.withFailureHandler(function(error){' +
              'message.className="error";' +
              'message.textContent=error&&error.message?error.message:"Operation failed.";' +
              'button.disabled=false;' +
              'button.textContent=' +
                JSON.stringify(buttonText) +
              ';' +
            '})' +
            '.submitApprovalEmailAction(' +
              JSON.stringify(data.token) +
              ',' +
              JSON.stringify(data.action) +
              ',remarks);' +
        '}' +
      '</script>' +
    '</body>' +
    '</html>'
  );
}


/**
 * Called from approval action HTML page.
 */
function submitApprovalEmailAction(
  token,
  action,
  remarks
) {
  const cleanToken =
    normalizeText_(token);

  const cleanAction =
    normalizeLower_(action);

  const cleanRemarks =
    normalizeText_(remarks)
      .substring(
        0,
        EMAIL_SERVICE_CONFIG
          .MAX_REMARK_LENGTH
      );

  if (
    cleanAction ===
    EMAIL_SERVICE_CONFIG
      .ACTION.APPROVE
  ) {
    return approveRequestByToken(
      cleanToken,
      cleanRemarks
    );
  }

  if (
    cleanAction ===
    EMAIL_SERVICE_CONFIG
      .ACTION.REJECT
  ) {
    if (!cleanRemarks) {
      return errorResponse_(
        new Error(
          'Rejection reason is required.'
        )
      );
    }

    return rejectRequestByToken(
      cleanToken,
      cleanRemarks
    );
  }

  return errorResponse_(
    new Error(
      'Invalid approval action.'
    )
  );
}


/**
 * Builds a simple result page.
 */
function buildApprovalMessagePage_(
  title,
  message,
  isSuccess
) {
  const color = isSuccess
    ? '#166534'
    : '#991b1b';

  const background = isSuccess
    ? '#dcfce7'
    : '#fee2e2';

  const html =
    '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
      '<meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<style>' +
        'body{font-family:Arial,sans-serif;background:#f3f4f6;margin:0;padding:20px;}' +
        '.box{max-width:560px;margin:60px auto;padding:25px;border-radius:12px;background:' +
          background +
        ';color:' +
          color +
        ';}' +
      '</style>' +
    '</head>' +
    '<body>' +
      '<div class="box">' +
        '<h2>' +
          escapeHtml_(title) +
        '</h2>' +
        '<p>' +
          escapeHtml_(message) +
        '</p>' +
      '</div>' +
    '</body>' +
    '</html>';

  return HtmlService
    .createHtmlOutput(html)
    .setTitle(title);
}


/**
 * Sends a test email using the latest pending request.
 *
 * First create a fresh pending request, then run:
 * testSendLatestApprovalEmail
 */
function testSendLatestApprovalEmail() {
  const pending =
    getPendingApprovalRequests();

  if (
    !pending.success ||
    !pending.data.records.length
  ) {
    throw new Error(
      'No pending approval request found. Create a new pending request first.'
    );
  }

  const requestId =
    pending.data.records[0]
      .requestId;

  const result =
    resendApprovalRequestEmail(
      requestId
    );

  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}


/**
 * Shows current email service configuration.
 */
function testEmailServiceConfig() {
  let webAppUrl = '';

  try {
    webAppUrl =
      getApprovalWebAppUrl_();
  } catch (error) {
    webAppUrl = '';
  }

  const result = {
    appName:
      EMAIL_SERVICE_CONFIG.APP_NAME,

    currentUserEmail:
      getCurrentUserEmail_(),

    webAppUrl:
      webAppUrl,

    webAppConfigured:
      Boolean(webAppUrl),

    remainingDailyQuota:
      MailApp.getRemainingDailyQuota()
  };

  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}



function saveProductionWebAppUrl() {
  return setApprovalWebAppUrl(
    'https://script.google.com/a/macros/anushagroup.com/s/AKfycbx1XfuvX43hb-kVy7MCwAmGp8Iwykpbu5BeQGBcb2TQ7DIBQYv01BWNEbBy_6qDOtQw9w/exec'
  );
}
