//Set this to true to send emails to the replyTo address instead of the one in the sheet (for testing)
const debugMode = false;
const productName = "On Demand Sandbox";
const eventName = "Request";
const sheetName = "Form Responses 1";
const actionLink = "https://docs.google.com/spreadsheets/d/1McZC5Brg6wLDCAIo7gEYjbyP20CrFfV8GK4d04nBsi4/edit?resourcekey#gid=634459441";
const playbookLink = "https://salesforce.quip.com/1SchArrxWUna";

//NOTE: approverEmailTo and approverEmailCc can be comma-delimited lists if you need multiple people involved for coverage / awareness
const approverEmailTo = "jonathan.tucker@salesforce.com";
const approverEmailCc = "";
const replyTo = "jonathan.tucker@salesforce.com";

  /*
  Assumed header text values in use:

  - Email Address
  - Partner company (it must be an existing SF partner)
  - Applicant full name
  - Region where you work (for statistical purposes)
  - Your partner email (preferably what you use to login into the Partner Community)
  - Who referred for this program? This is not open for every partner, so we need to know who started this process with you or your company
  - ODS
  - TTL
  - Sandbox ID
  - Notes

  Mapping all columns by index makes debugging Apps Script far easier along with getting older data when it might be needed later for responses to duplicates, etc.
  */

const emailAddressHeader = "Email Address";
const emailAddressColumnIndex = getHeaderIndex(sheetName, emailAddressHeader);
const partnerCompanyHeader = "Partner company (it must be an existing SF partner)";
const partnerCompanyColumnIndex = getHeaderIndex(sheetName, partnerCompanyHeader);
const applicantFullNameHeader = "Applicant full name";
const applicantFullNameColumnIndex = getHeaderIndex(sheetName, applicantFullNameHeader);
const regionHeader = "Region where you work (for statistical purposes)";
const regionColumnIndex = getHeaderIndex(sheetName, regionHeader);
const yourPartnerEmailHeader = "Your partner email (preferably what you use to login into the Partner Community)";
const yourPartnerEmailColumnIndex = getHeaderIndex(sheetName, yourPartnerEmailHeader);
const whoReferredHeader = "Who referred for this program?  This is not open for every partner, so we need to know who started this process with you or your company";
const whoReferredColumnIndex = getHeaderIndex(sheetName, whoReferredHeader);
const odsHeader = "ODS";
const odsColumnIndex = getHeaderIndex(sheetName, odsHeader);
const ttlHeader = "TTL";
const ttlColumnIndex = getHeaderIndex(sheetName, ttlHeader);
const sandboxIdHeader = "Sandbox ID";
const sandboxIdColumnIndex = getHeaderIndex(sheetName, sandboxIdHeader);
const notesHeader = "Notes";
const notesColumnIndex = getHeaderIndex(sheetName, notesHeader);

function onFormSubmit(e)
{
  //Get the row of submitted data by names and other means where needed
  let submission = { 
    emailAddress: e.values[emailAddressColumnIndex]
    , partnerCompany: e.values[partnerCompanyColumnIndex]    
    , applicantFullName: e.values[applicantFullNameColumnIndex]
    , region: e.values[regionColumnIndex]
    , yourPartnerEmail: e.values[yourPartnerEmailColumnIndex]
    , whoReferred: e.values[whoReferredColumnIndex]
  };
  
  //You can avoid mistakes or spamming when first assigning this code to your form using the debugMode variable at the top of the code
  if(debugMode)
  {
    submission.yourPartnerEmail = replyTo;
  }
  
  //Check if we have a record on file for the email address and same session time
  Logger.log("submission.emailAddress: " + submission.emailAddress);
  Logger.log("submission.opportunityCustomerName: " + submission.yourPartnerEmail);
  let alreadyRegisteredResult = isDuplicateRegistration(submission.yourPartnerEmail, submission.partnerCompany);
  
  if(alreadyRegisteredResult.alreadyRegistered === true)
  {
    //Duplicate request email to the person registering
    MailApp.sendEmail({
      to: submission.yourPartnerEmail
      , replyTo: replyTo
      , subject: productName + " " + eventName + " (Duplicate)"
      , htmlBody: "Dear " + submission.applicantFullName + ",<br /><br /> Thanks for submitting your <b>" + productName + " " + eventName + "</b>. "
      + "Please note that this looks like a duplicate request and it will be reviewed for approval as soon as possible if not already provided to you.<br /><br />"
      + "Please also note that this form is not meant for extending On Demand Sandboxes. As a matter of policy, we do not provide extensions to an ODS. You can back up the contents of the sandbox per <a href=\"https://documentation.b2c.commercecloud.salesforce.com/DOC1/topic/com.demandware.dochelp/content/b2c_commerce/topics/import_export/b2c_using_site_import_export_for_development_testing.html\">this link</a>. If you need a sandbox for a longer period, you should request an ODS from your customer.<br /><br />"
      + makeRegistrationTable(submission) + "<br />"
      + "Please allow 3 business days for review and provisioning.<br />"
    });
    Logger.log("Sent email for duplicate request: " + productName + " " + eventName + " to <" + submission.yourPartnerEmail + ">");
    
    //Duplicate request notification to approver(s)
    MailApp.sendEmail({
      to: approverEmailTo
      , cc: approverEmailCc
      , replyTo: replyTo
      , subject: productName + " " + eventName + " (Duplicate)"
      , htmlBody: submission.applicantFullName + " submitted a duplicate <b>" + productName + " " + eventName + "</b>. Review may be needed:<br /><br />"
      + makeRegistrationTable(submission) + "<br /><br />"
      + "<a href=\"" + actionLink + "\">Click here</a> to review and action the request or <a href=\"" + playbookLink + "\">click here</a> for the playbook.<br />"
    });
  }
  else
  {
    //NOTE: If your form allows editing, you will get an exception for not having a recipient here when editing - so don't  allow editing on your form settings ;)
    
    //Initial request email to the person registering
      MailApp.sendEmail({
      to: submission.yourPartnerEmail
      , replyTo: replyTo
      , subject: productName + " " + eventName + " (Initial)"
      , htmlBody: "Dear " + submission.applicantFullName + ",<br /><br /> Thanks for submitting your <b>" + productName + " " + eventName + "</b>. "
      + "We have the following data on file:<br /><br />"
      + makeRegistrationTable(submission) + "<br />"
      + "Please allow 3 business days for review and provisioning.<br />"
    });

    //Initial request email
      MailApp.sendEmail({
      to: approverEmailTo
      , cc: approverEmailCc
      , replyTo: replyTo
      , subject: "Action Required: " + productName + " " + eventName + " Review (New)"
      , htmlBody: submission.applicantFullName + " submitted a new <b>" + productName + " " + eventName + "</b>. <u>Review and actioning is needed</u>:<br /><br />"
      + makeRegistrationTable(submission) + "<br /><br />"
      + "<a href=\"" + actionLink + "\">Click here</a> to review and action the request or <a href=\"" + playbookLink + "\">click here</a> for the playbook.<br />"
    });

    Logger.log("Sent email for initial request: " + productName + " " + eventName + " to <" + submission.emailAddress + ">");
  }
}

function getHeaderIndex(sheetName, headerText)
{
  let result = -1;
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  if(sheet == null)
  {
    console.log('Sheet not located in getIndexByHeader using: ' + sheetName);
    return result;
  }

  let data = sheet.getDataRange().getValues();

  for (let i=0; i < data[0].length; i++)
  {
    if(data[0][i].toString() == headerText)
    {
      result = i;
      break;
    }
  }
  return result;
}

//Make a fungible old-school HTML table of name-value pairs
function makeRegistrationTable(data)
{
  let registrationTable = "";
  registrationTable += "<b>Request Type</b>: " + productName + " " + eventName  + "<br />";
  registrationTable += "<b>" + "Partner Company" + "</b>: " + data.partnerCompany + "<br />";
  registrationTable += "<b>" + applicantFullNameHeader + "</b>: " + data.applicantFullName + "<br />";
  registrationTable += "<b>" + "Region" + "</b>: " + data.region + "<br />";
  registrationTable += "<b>" + "Partner Email Address" + "</b>: " + data.yourPartnerEmail + "<br />";
  registrationTable += "<b>" + "Referred by" + "</b>: " + data.whoReferred + "<br />";
  return registrationTable;
}

//Detect duplicates based on composite of the email address and training slot
function isDuplicateRegistration(yourPartnerEmail, partnerCompany)
{
  Logger.log("isDuplicateRegistration(" + yourPartnerEmail + "," + partnerCompany + ") invoked.");
  //The current submission does not get stopped so we need to just track it and kill the real dup looking at any matches added beyond a length of 1
  let hits = [];
  let sheet = SpreadsheetApp.getActiveSheet();
  let data  = sheet.getDataRange().getValues();
  let i = 0;
        
  for (i = 0; i < data.length; i++)
  {
    if (data[i][yourPartnerEmailColumnIndex].toString().toLowerCase() === yourPartnerEmail.toLowerCase()
    && data[i][partnerCompanyColumnIndex].toString().toLowerCase() === partnerCompany.toLowerCase())
    {
      Logger.log("Pushing: " + JSON.stringify(data[i]));  
      hits.push(data[i]);
    }
  }
  
  //Return some enriched results we can use
  if(hits.length > 1)
  {
    Logger.log("Located duplicate partner email: <" + yourPartnerEmail + "> with Company: " + partnerCompany.toLowerCase());
    //Remove the newest addition using the tracking array (like it never happened)
    sheet.deleteRow(i);
    //Get the data for the original (old) submission into something we can use
    let originalRow = hits[hits.length - 2];
    
    let priorRegistration = {
      emailAddress: originalRow[emailAddressColumnIndex]
    , partnerCompany: originalRow[partnerCompanyColumnIndex]    
    , applicantFullName: originalRow[applicantFullNameColumnIndex]
    , region: originalRow[regionColumnIndex]
    , yourPartnerEmail: originalRow[yourPartnerEmailColumnIndex]
    , whoReferred: originalRow[whoReferredColumnIndex] 
    };

    Logger.log("Located duplicate partner email: <" + yourPartnerEmail + "> with Company: " + partnerCompany.toLowerCase());
    return JSON.parse("{\"alreadyRegistered\": true, \"priorData\": " + JSON.stringify(priorRegistration) +"}");
  }
  else
  {
    Logger.log("Did not find duplicate email: <" + yourPartnerEmail + "> with Company: " + partnerCompany.toLowerCase());
    return JSON.parse("{\"alreadyRegistered\": false, \"priorData\": null}");
  }
}
