/**
 * 0) read the README to understand how this works
 * 1) Change these to match the column names you are using for email 
 * recepient addresses, email sent column, and days since reply.
 * Our convention was to label all initial draft subjects as follows:
 * "Berkeley Consulting Club Seeks {company name} as a Client!",
 * where company name was filled using the sheet info, and followups as:
 * "Followup Template FA21".
 * 2) Change the draft and phrase variables to match your desired 
 * convention. The alterations of firstDraft and nextPhrase just add 
 * spaces (either before or after) to assist with string manipulation
 * 3) change the sheet ID to match yours (found in url)
 * Our schedule send utilized cell J2 to determine the last row, 
 * 4) change to your desired cell
 * 
**/
const RECIPIENT1_COL  = "Recipient";
const EMAIL_SENT_COL = "Email Sent";
const INDIC = "Days Since Reply";
const firstDraft = "Berkeley Consulting Club Seeks";
const firstDraftSpace = "Berkeley Consulting Club Seeks ";
const nextPhraseExact = "as a client!";
const nextPhraseSearch = " as a client";
const lastRowInfo = "J2";
const sheetID = "1TDhcmI3KjUFo6RXeSKNmMpfO_ZZJ8wxYrzTJ89o94jY";
const followupDraftSubject = "Followup Template FA21";
 
/** 
 * Creates the menu item "Mail Merge" for user to run scripts on drop-down.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
      .addItem('Send Emails', 'sendEmails')
      .addToUi();
}
 
/**
 * Send emails from sheet data.
 * @param {string} subjectLine (optional) for the email draft message
 * @param {Sheet} sheet to read data from
*/
function sendEmails() {
  // option to skip browser prompt if you want to use this code in other projects
  const ss = SpreadsheetApp.openById(sheetID)
  const sheet = SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
  
  lastRow = Browser.inputBox("Mail Merge", "enter the last row you want the auto-send to consider", Browser.Buttons.OK_CANCEL);
 
  if (!lastRow) {
    lastRow = Number(sheet.getRange(lastRowInfo).getValue());
  }

  const emailTemplate = getGmailTemplateFromDrafts_(firstDraft);
  
  // get the data from the passed sheet
  const dataRange = sheet.getRange(1, 1, lastRow, 9);

  // Fetch displayed values for each row in the Range H
  
  const data = dataRange.getDisplayValues();

  // assuming row 1 contains our column headings
  const heads = data.shift(); 
  
  // get the index of column named 'Email Status' (Assume header names are unique)
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);
  
  // convert 2d array into object array
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  // used to record sent emails
  const out = [];

  // loop through all the rows of data
  obj.forEach(function(row, rowIdx){
    // only send emails is email_sent cell is blank and not hidden by filter
    if (row[EMAIL_SENT_COL] == ''){
      try {
        const tot = fillInTemplateFromObject_(emailTemplate.message, row);
        const msgObj = tot[0];
        const companyName = tot[1];

        // handles first emails
        if (tot[2] == "No") {
          GmailApp.sendEmail(row[RECIPIENT1_COL], msgObj.subject + " " + companyName + " " + nextPhraseExact, msgObj.text, {
            htmlBody: msgObj.html,
            // bcc: 'a.bbc@email.com',
            // cc: 'a.cc@email.com',
            // from: 'an.alias@email.com',
            // name: 'Shikhar Srivastava',
            // replyTo: 'a.reply@email.com',
            // noReply: true, // if the email should be sent from a generic n
          });
          // modify cell to record email sent date
          Utilities.sleep(10000);
          var latestEmails = GmailApp.getInboxThreads(0, 30);
          var bounce = false;
          for (var i = 0; i < latestEmails.length; i++) {
            var tempMsgs = latestEmails[i].getMessages();
            var tempEmail = tempMsgs[tempMsgs.length - 1];
            var tempOrig = tempMsgs[0];
            if (tempEmail.getFrom() == "Mail Delivery Subsystem <mailer-daemon@googlemail.com>" && tempOrig.getTo() == row[RECIPIENT1_COL]) { 
              bounce = true;
            }
          }
          if (bounce) {
            out.push(["Email Bounced"]);
          } else {
            out.push([Utilities.formatDate(new Date(), "PST", "MM-dd-yyyy")]);
          }
        } else {
          // handles follow up emails
          var replyThread = GmailApp.search(firstDraftSpace + companyName + nextPhraseSearch);
          for (var j = 0; j < replyThread.length; j++) {
            if (replyThread[j].getMessages()[0].getTo().includes(row[RECIPIENT1_COL])) {
              var realThread = replyThread[j];
              break;
            }
          }
          recP = realThread.getMessages()[0].getTo();
          var followUpT = getGmailTemplateFromDrafts_(followupDraftSubject);
          var toSend = fillInTemplateFromObject_(followUpT.message, row)[0];
          var draft = realThread.createDraftReply(toSend.text, {htmlBody: toSend.html});
          const rawMsg = Gmail.Users.Drafts.get("me", draft.getId(), {format: "raw"}).message;
          const msgString = rawMsg.raw.reduce(function (acc, b) { return acc + String.fromCharCode(b);}, "");
          const pattern = /^To: .+$/m;
          var new_msg_string = msgString.replace(pattern, "To: <" + recP + ">");
          const encoded_msg = Utilities.base64EncodeWebSafe(new_msg_string);

          const resource = {
            id: draft.getId(),
            message: {
              threadId: realThread.getId(),
              raw: encoded_msg
            }
          }
          const resp = Gmail.Users.Drafts.update(resource, "me", draft.getId());
          Gmail.Users.Drafts.send({id: resp.id}, "me");
          out.push([Utilities.formatDate(new Date(), "PST", "MM-dd-yyyy")]);
        }
      } catch(e) {
        // modify cell to record error
        out.push([e.message]);
      }
    } else {
      out.push([row[EMAIL_SENT_COL]]);
    }
  });
  
  // updating the sheet with new data
  sheet.getRange(2, emailSentColIdx+1, out.length).setValues(out);
  
  /**
   * Get a Gmail draft message by matching the subject line.
   * @param {string} subject_line to search for draft message
   * @return {object} containing the subject, plain and html message body and attachments
  */
  function getGmailTemplateFromDrafts_(subject_line){
    Logger.log(subject_line);
    try {
      // get drafts
      const drafts = GmailApp.getDrafts();
      // filter the drafts that match subject line
      const draft = drafts.filter(subjectFilter_(subject_line))[0];
      // get the message object
      const msg = draft.getMessage();
      // getting attachments so they can be included in the merge
      const attachments = msg.getAttachments();
      return {message: {subject: subject_line, text: msg.getPlainBody(), html:msg.getBody()}, 
              attachments: attachments};
    } catch(e) {
      Logger.log(e);
      throw new Error("Oops - can't find Gmail draft");
    }

    /**
     * Filter draft objects with the matching subject linemessage by matching the subject line.
     * @param {string} subject_line to search for draft message
     * @return {object} GmailDraft object
    */
    function subjectFilter_(subject_line){
      return function(element) {
        if (element.getMessage().getSubject() === subject_line) {
          return element;
        }
      }
    }
  }
  
  /**
   * Fill template string with data object
   * @see https://stackoverflow.com/a/378000/1027723
   * @param {string} template string containing {{}} markers which are replaced with data
   * @param {object} data object used to replace {{}} markers
   * @return {object} message replaced with data
  */
  function fillInTemplateFromObject_(template, data) {
    // we have two templates one for plain text and the html body
    // stringifing the object means we can do a global replace
    let template_string = JSON.stringify(template);

    const compName = data["CompanyName"];
    const followUpIndicator = data["FollowUp?"];
    // token replacement
    template_string = template_string.replace(/{{[^{}]+}}/g, key => {
      return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
    });
    
    return [JSON.parse(template_string), compName, followUpIndicator];
  }

  /**
   * Escape cell data to make JSON safe
   * @see https://stackoverflow.com/a/9204218/1027723
   * @param {string} str to escape JSON special characters from
   * @return {string} escaped string
  */
  function escapeData_(str) {
    return str
      .replace(/[\\]/g, '\\\\')
      .replace(/[\"]/g, '\\\"')
      .replace(/[\/]/g, '\\/')
      .replace(/[\b]/g, '\\b')
      .replace(/[\f]/g, '\\f')
      .replace(/[\n]/g, '\\n')
      .replace(/[\r]/g, '\\r')
      .replace(/[\t]/g, '\\t');
  };
}
// function for refreshing the status of 'days since reply'
function refreshDates() {
  const ss = SpreadsheetApp.openById(sheetID)
  const sheet = SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
  const dataRange = sheet.getRange(2, 1, 307, 9);
  const dataInput = dataRange.getDisplayValues();
  const dataOutput = [];
  var b=0;
  for (var i=0; i < dataInput.length; i++) {
    var oldDate = dataInput[i][7];
    if (!(oldDate == "") && !(oldDate == "skip")) {
      var potThreads = GmailApp.search(firstDraftSpace + dataInput[i][4] + nextPhraseSearch);
      for (var j = 0; j < potThreads.length; j++) {
            if (potThreads[j].getMessages()[0].getTo().includes(dataInput[i][0])) {
              var realThread = potThreads[j];
              break;
            }
          }

      Logger.log(realThread.getLastMessageDate());
      Logger.log(firstDraftSpace + dataInput[i][4] + nextPhraseSearch);
      newDate = Utilities.formatDate(realThread.getLastMessageDate(), "PST", "MM-dd-yyyy");
      dataOutput[b] = []
      dataOutput[b].push(newDate);
      b++;
    } else {
      dataOutput[b] = [];
      dataOutput[b].push(oldDate);
      b++;
    }
  }
  var outRange = sheet.getRange(2, 8, 307);
  outRange.setValues(dataOutput);
}