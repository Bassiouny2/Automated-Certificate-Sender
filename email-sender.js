function sendCertificates() {
  // Constants
  const SUBJECT = "Your Participation Pass is ready"; // Email subject line
  const FROM_NAME = "UAE Hackathon"; // Sender's name
  const HTML_TEMPLATE = "EmailTemplate"; // HTML template file name for the email body
  const PPT_ID = "1gvMvkp1Vi1fN7oIM5wUc3aTkSqRaB7WgRcCgXGRXNuQ"; // Google Drive ID of the PowerPoint template

  var sheet = SpreadsheetApp.getActiveSheet(); // Get the active sheet
  var dataRange = sheet.getRange("A2:C" + sheet.getLastRow()); // Get the range of data from column A to C, starting from row 2
  var data = dataRange.getValues(); // Get all values in the specified range
  var emailTemplate =
    HtmlService.createHtmlOutputFromFile(HTML_TEMPLATE).getContent(); // Get the content of the email template

  var sentCount = 0; // Counter for sent emails

  // Assuming data (FirstName, LastName, EmailAddress, isSent)
  for (var i = 0; i < data.length; i++) {
    var Name = data[i][0]; // Participant's full name
    var recipientEmail = data[i][1]; // Participant's email address
    var isSent = data[i][2]; // Status of whether the certificate has been sent

    var names = Name.split(" ");
    Name = names.slice(0, 3).join(" "); // Use the first three names if there are more
    var FName = names[0]; // Participant's first name

    if (!isSent) {
      var personalizedEmail = emailTemplate.replace("{{FName}}", FName); // Personalize the email template
      var pptFile = DriveApp.getFileById(PPT_ID); // Get the PowerPoint template file
      var pptCopy = pptFile.makeCopy(); // Make a copy of the template
      var pptPresentation = SlidesApp.openById(pptCopy.getId()); // Open the copied template

      var slides = pptPresentation.getSlides(); // Get all slides in the presentation

      // Loop over each slide (assuming multiple slides)
      for (var slideIndex = 0; slideIndex < slides.length; slideIndex++) {
        var slide = slides[slideIndex];

        var shapes = slide.getShapes(); // Get all shapes in the slide

        // Loop over all shapes in the slide
        for (var shapeIndex = 0; shapeIndex < shapes.length; shapeIndex++) {
          var textRange = shapes[shapeIndex].getText(); // Get the text in the shape
          var text = textRange.asString(); // Convert text to string

          // Replace the placeholders with the participant's name
          var newText = text.replace("{{Name}}", Name);

          textRange.clear(); // Clear the existing text
          textRange.insertText(0, newText); // Insert the new text
        }
      }

      // Save and close the presentation after modifications
      pptPresentation.saveAndClose();

      // Convert the presentation into a PDF blob
      var pdfBlob = DriveApp.getFileById(pptCopy.getId()).getBlob();
      var pdfFile = DriveApp.createFile(pdfBlob); // Create a PDF file from the blob
      pdfFile.setName("HackathonID_" + Name + ".pdf"); // Name the PDF file

      try {
        // Send the email with the PDF attachment
        GmailApp.sendEmail(recipientEmail, SUBJECT, "", {
          htmlBody: personalizedEmail,
          attachments: [pdfFile.getAs(MimeType.PDF)],
          name: FROM_NAME,
        });

        data[i][2] = true; // Update the "isSent" status to true
        sentCount++;
        dataRange.setValues(data); // Update the "isSent" column values in the spreadsheet after processing the batch

        pptCopy.setTrashed(true); // Delete the temporary PPT copy
        pdfFile.setTrashed(true); // Delete the temporary PDF file

        Logger.log(
          "Certificate sent to " + recipientEmail + ". Sent Count: " + sentCount
        );

        if (i % 5 === 0) {
          Utilities.sleep(1000);
        } // Sleep for 1 second every 5 emails
        if (sentCount >= 1999) break; // Stop if the sent count reaches 1999
      } catch (error) {
        Logger.log("Error sending certificate: " + error); // Log any errors
      }
    }
  }
}
