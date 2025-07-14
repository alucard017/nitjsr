const coordinators = {
  'Shubham': {
    name: "Shubham Mundhra",
    phone: "+91 1234567890",
    email: "shubham@gmail.com"
  },
  'Shwetabh': {
    name: "Shwetabh Sinha",
    phone: "+91 1234567890",
    email: "shwetabh@gmail.com
  },
  'Priya': {
    name: "Priya Surendra Tiwari",
    phone: "+91 1234567890",
    email: "priya@gmail.com"
  }
};

const ATTACHMENT_IDS = [
  'BROCHURE_ID',
  'JAF_ID',
  'IAF_ID'
];

function generateEmailBody(companyName, hrName, sentBy) {
  const senderContact = coordinators[sentBy];
  if (!senderContact) throw new Error(`Coordinator ${sentBy} not found.`);

  // Sort all coordinators with sender first
  const allContacts = Object.values(coordinators);
  const sortedContacts = [senderContact, ...allContacts.filter(c => c.name !== senderContact.name)];

  const template = HtmlService.createTemplateFromFile('htmlTemplate');
  template.CompanyName = companyName;
  template.hrName = hrName;
  template.CoordinatorName = senderContact.name;
  template.CoordinatorPhone = senderContact.phone;
  template.CoordinatorEmail = senderContact.email;
  template.contacts = sortedContacts; // for looping in the template

  return template.evaluate().getContent();
}

function sendMailsDirectly() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const [companyName, hrName, recipient, cc, sentBy, emailSent] = data[i];

   if (!emailSent || emailSent.toString().trim() === "") {
    const subject = `Campus Placement & Internship Invitation - NIT Jamshedpur`;
    const htmlBody = generateEmailBody(companyName, hrName, sentBy);


      // Measure HTML body size
      const htmlSize = htmlBody.length;

      // Get attachment blobs and measure total size
      let attachmentsSize = 0;
      const attachments = ATTACHMENT_IDS.map(fileId => {
        const file = DriveApp.getFileById(fileId);
        const blob = file.getMimeType() === MimeType.GOOGLE_DOCS
          ? file.getAs(MimeType.PDF)
          : file.getBlob();
        attachmentsSize += blob.getBytes().length;
        return blob;
      });

      const totalSizeBytes = htmlSize + attachmentsSize;
      const totalSizeKB = totalSizeBytes / 1024;
      const totalSizeMB = totalSizeKB / 1024;

      Logger.log(`Email to ${recipient} | Size: ${totalSizeKB.toFixed(2)} KB (${totalSizeMB.toFixed(2)} MB)`);

      GmailApp.sendEmail(recipient, subject, "", {
        htmlBody,
        cc,
        attachments,
        name: coordinators[sentBy]?.name || 'NIT Jamshedpur'
      });

      // Labeling
      // const fullLabel = `MCA 2K26 Batch/${sentBy}`;
      // let label = GmailApp.getUserLabelByName(fullLabel);
      // if (!label) label = GmailApp.createLabel(fullLabel);

      // const threads = GmailApp.search(`to:${recipient} subject:"${subject}" newer_than:1d`);
      // if (threads.length > 0) {
      //   threads[0].addLabel(label);
      // }
      const parentLabelName = `MCA 2K26 Batch`;
      const childLabelName = `${parentLabelName}/${sentBy}`;

      let parentLabel = GmailApp.getUserLabelByName(parentLabelName);
      if (!parentLabel) parentLabel = GmailApp.createLabel(parentLabelName);

      let childLabel = GmailApp.getUserLabelByName(childLabelName);
      if (!childLabel) childLabel = GmailApp.createLabel(childLabelName);

      const threads = GmailApp.search(`to:${recipient} subject:"${subject}" newer_than:1d`);
      if (threads.length > 0) {
        threads[0].addLabel(parentLabel);
        threads[0].addLabel(childLabel);
      }
      // Set sent date in the sheet (Column F = Date)
      sheet.getRange(i + 1, 6).setValue(new Date());
    }
  }
}

// Add Parent Labels to all sent unlabelled mails
function addParentLabelToThreads() {
  const parentLabelName = 'MCA 2K26 Batch';
  const parentLabel = GmailApp.getUserLabelByName(parentLabelName) 
                      || GmailApp.createLabel(parentLabelName);

  const subLabelPrefix = `${parentLabelName}/`;
  const allLabels = GmailApp.getUserLabels();

  allLabels.forEach(label => {
    const name = label.getName();
    if (name.startsWith(subLabelPrefix)) {
      const threads = label.getThreads();

      threads.forEach(thread => {
        const threadLabels = thread.getLabels();
        const hasParentLabel = threadLabels.some(l => l.getName() === parentLabelName);

        if (!hasParentLabel) {
          thread.addLabel(parentLabel);
          Logger.log(`Added parent label to thread with subject: ${thread.getFirstMessageSubject()}`);
        }
      });
    }
  });

  Logger.log("Parent label applied to all applicable threads.");
}
