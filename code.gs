let MySheets  = SpreadsheetApp.getActiveSpreadsheet();
let LoginSheet  = MySheets.getSheetByName("login"); 

function doGet(e) {
  var output = HtmlService.createTemplateFromFile('login');
  
  var sess = getSession();
   if (sess.loggedIn) {
     output = HtmlService.createTemplateFromFile('main');
  }
  return output.evaluate();
}

function myURL() {
  return ScriptApp.getService().getUrl();
}


function setSession(session) {
  var sId   = Session.getTemporaryActiveUserKey();
  var uProp = PropertiesService.getUserProperties();
  uProp.setProperty(sId, JSON.stringify(session));
}


function getSession() {
  var sId   = Session.getTemporaryActiveUserKey();
  var uProp = PropertiesService.getUserProperties();
  var sData = uProp.getProperty(sId);
  return sData ? JSON.parse(sData) : { loggedIn: false };
}


function loginUser(pUID, pPassword) {
    
    if (loginCheck(pUID, pPassword)) {
      
      var sess = getSession();
      sess.loggedIn = true;
      setSession(sess);

        return 'success';
    } 
    else {
        return 'failure';
    }
}


function logoutUser() {
  var sess = getSession();
  sess.loggedIn = false;
  setSession(sess);
}


function loginCheck(pUID, pPassword) {
  let LoginPass =  false;
      let ReturnData = LoginSheet.getRange("A:A").createTextFinder(pUID).matchEntireCell(true).findAll();
        
        ReturnData.forEach(function (range) {
          let StartRow = range.getRow();
          let TmpPass = LoginSheet.getRange(StartRow, 2).getValue();
          if (TmpPass == pPassword)
          {
              LoginPass = true;
          }
        });

    return LoginPass;
}

function OpenPage(PageName)
{
    return HtmlService.createHtmlOutputFromFile(PageName).getContent();
}


function UserRegister(pUID, pPassword, pName) {
    
    let RetMsg = '';
    let ReturnData = LoginSheet.getRange("A:A").createTextFinder(pUID).matchEntireCell(true).findAll();
    let StartRow = 0;
    ReturnData.forEach(function (range) {
      StartRow = range.getRow();
    });

    if (StartRow > 0) 
    {
      RetMsg = 'danger, User Already Exists';
    }
    else
    {
      LoginSheet.appendRow([pUID, pPassword, pName]) ;  
      RetMsg = 'success, User Successfully Registered'; 
    }

    return  RetMsg;
}
// Templates
var templates = {
  "Template1": { id: 'replace with template 1 id' },
  "Template2": { id: 'replace with template 2 id' },
  "Template3": { id: 'replace with template 3 id' }
};

// Function to update the live preview
function getUpdatedPreview(formData, selectedTemplate) {
  const templateId = templates[selectedTemplate].id;
  const doc = DocumentApp.openById(templateId).getBody();
  let templateText = doc.getText();

  // Replace placeholders with form data
  templateText = templateText
    .replace('{{firstname}}', formData.firstname || 'Your First Name')
    .replace('{{lastname}}', formData.lastname || 'Your Last Name')
    .replace('{{designation}}', formData.designation || 'Your Designation')
    .replace('{{address}}', formData.address || 'Your Address')
    .replace('{{email}}', formData.email || 'your.email@example.com')
    .replace('{{phoneno}}', formData.phoneno || '123-456-7890')
    .replace('{{summary}}', formData.summary || 'A brief summary about yourself.');

  // Handle arrays (e.g., achievements, experience, etc.)
  if (formData.achievements && formData.achievements.length > 0) {
    templateText = templateText.replace('{{achieve_title}}', formData.achievements.map(a => a.title).join(', '));
    templateText = templateText.replace('{{achieve_description}}', formData.achievements.map(a => a.description).join(', '));
  }

  if (formData.experience && formData.experience.length > 0) {
    templateText = templateText.replace('{{exp_title}}', formData.experience.map(e => e.title).join(', '));
    templateText = templateText.replace('{{exp_organization}}', formData.experience.map(e => e.organization).join(', '));
    templateText = templateText.replace('{{exp_location}}', formData.experience.map(e => e.location).join(', '));
    templateText = templateText.replace('{{exp_start_date}}', formData.experience.map(e => e.startDate).join(', '));
    templateText = templateText.replace('{{exp_end_date}}', formData.experience.map(e => e.endDate).join(', '));
    templateText = templateText.replace('{{exp_description}}', formData.experience.map(e => e.description).join(', '));
  }

  if (formData.education && formData.education.length > 0) {
    templateText = templateText.replace('{{edu_school}}', formData.education.map(edu => edu.school).join(', '));
    templateText = templateText.replace('{{edu_degree}}', formData.education.map(edu => edu.degree).join(', '));
    templateText = templateText.replace('{{edu_city}}', formData.education.map(edu => edu.city).join(', '));
    templateText = templateText.replace('{{edu_start_date}}', formData.education.map(edu => edu.startDate).join(', '));
    templateText = templateText.replace('{{edu_graduation_date}}', formData.education.map(edu => edu.graduationDate).join(', '));
    templateText = templateText.replace('{{edu_description}}', formData.education.map(edu => edu.description).join(', '));
  }

  if (formData.projects && formData.projects.length > 0) {
    templateText = templateText.replace('{{proj_title}}', formData.projects.map(p => p.title).join(', '));
    templateText = templateText.replace('{{proj_link}}', formData.projects.map(p => p.link).join(', '));
    templateText = templateText.replace('{{proj_description}}', formData.projects.map(p => p.description).join(', '));
  }

  if (formData.skills && formData.skills.length > 0) {
    templateText = templateText.replace('{{skills}}', formData.skills.join(', '));
  }

  return templateText.replace(/\n/g, '<br>'); // Return HTML-friendly text
}
function processForm(formData, selectedTemplate) {
  const folderId = 'replace with folder id'; // Replace with your actual Google Drive Folder ID
  const templateId = templates[selectedTemplate].id;
  
  // Create a copy of the selected template
  const newFile = DriveApp.getFileById(templateId).makeCopy(formData.firstname + ' ' + formData.lastname, DriveApp.getFolderById(folderId));
  const doc = DocumentApp.openById(newFile.getId());
  const body = doc.getBody();

  // Function to get default value if data is empty
  function getDefault(value, defaultValue) {
    return value && value.trim() !== '' ? value : defaultValue;
  }

  // Replace text placeholders in the document with default values
  body.replaceText('{{firstname}}', getDefault(formData.firstname, 'Your First Name'));
  body.replaceText('{{lastname}}', getDefault(formData.lastname, ' '));
  body.replaceText('{{designation}}', getDefault(formData.designation, 'Your Designation'));
  body.replaceText('{{address}}', getDefault(formData.address, 'Your Address'));
  body.replaceText('{{email}}', getDefault(formData.email, 'your.email@example.com'));
  body.replaceText('{{phoneno}}', getDefault(formData.phoneno, '123-456-7890'));
  body.replaceText('{{summary}}', getDefault(formData.summary, 'A brief summary about yourself.'));

  // Insert Profile Image at {{profile_image}} Placeholder
  if (formData.image && formData.image.trim() !== '') {
  try {
    const imgBlob = Utilities.newBlob(Utilities.base64Decode(formData.image.split(",")[1]), "image/png", "profile.png");
    const imgFile = DriveApp.getFolderById(folderId).createFile(imgBlob);
    const img = imgFile.getAs("image/png");

    const rangeElement = body.findText("{{profile_image}}");
    if (rangeElement) {
      const startOffset = rangeElement.getStartOffset();
      const paragraph = rangeElement.getElement().getParent(); // Get the paragraph containing the placeholder

      paragraph.replaceText("{{profile_image}}", ""); // Remove placeholder
      paragraph.insertInlineImage(startOffset, img).setWidth(100).setHeight(100); // Adjust size as needed
    }
  } catch (error) {
    Logger.log("Error inserting image: " + error.message);
  }
} else {
  // Remove the placeholder if no image is provided
  body.replaceText("{{profile_image}}", "");
}



  // Replace Achievements
  body.replaceText('{{achieve_title}}', getDefault(formData.achievements?.map(a => a.title).join(', '), 'Your Achievements'));
  body.replaceText('{{achieve_description}}', getDefault(formData.achievements?.map(a => a.description).join(', '), 'Achievement details'));

  // Replace Experience
  body.replaceText('{{exp_title}}', getDefault(formData.experience?.map(e => e.title).join(', '), 'Your Job Title'));
  body.replaceText('{{exp_organization}}', getDefault(formData.experience?.map(e => e.organization).join(', '), 'Company/Organization'));
  body.replaceText('{{exp_location}}', getDefault(formData.experience?.map(e => e.location).join(', '), 'Location'));
  body.replaceText('{{exp_start_date}}', getDefault(formData.experience?.map(e => e.startDate).join(', '), 'Start Date'));
  body.replaceText('{{exp_end_date}}', getDefault(formData.experience?.map(e => e.endDate).join(', '), 'End Date'));
  body.replaceText('{{exp_description}}', getDefault(formData.experience?.map(e => e.description).join(', '), 'Job Description'));

  // Replace Education
  body.replaceText('{{edu_school}}', getDefault(formData.education?.map(edu => edu.school).join(', '), 'Your School/University'));
  body.replaceText('{{edu_degree}}', getDefault(formData.education?.map(edu => edu.degree).join(', '), 'Your Degree'));
  body.replaceText('{{edu_city}}', getDefault(formData.education?.map(edu => edu.city).join(', '), 'City'));
  body.replaceText('{{edu_start_date}}', getDefault(formData.education?.map(edu => edu.startDate).join(', '), 'Start Date'));
  body.replaceText('{{edu_graduation_date}}', getDefault(formData.education?.map(edu => edu.graduationDate).join(', '), 'Graduation Date'));
  body.replaceText('{{edu_description}}', getDefault(formData.education?.map(edu => edu.description).join(', '), 'Education Details'));

  // Replace Projects
  body.replaceText('{{proj_title}}', getDefault(formData.projects?.map(p => p.title).join(', '), 'Project Title'));
  body.replaceText('{{proj_link}}', getDefault(formData.projects?.map(p => p.link).join(', '), 'Project Link'));
  body.replaceText('{{proj_description}}', getDefault(formData.projects?.map(p => p.description).join(', '), 'Project Description'));

  // Replace Skills
  body.replaceText('{{skills}}', getDefault(formData.skills?.join(', '), 'Your Skills'));

  doc.saveAndClose();

  // Convert to PDF
  const pdfFile = newFile.getAs('application/pdf');
  const pdf = DriveApp.getFolderById(folderId).createFile(pdfFile);

  // Delete the temporary Google Doc
  newFile.setTrashed(true);

  // Set sharing permissions for the PDF and return the link
  pdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return `https://drive.google.com/uc?export=download&id=${pdf.getId()}`;
}
