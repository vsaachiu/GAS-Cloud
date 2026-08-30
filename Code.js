var thisUser = Session.getActiveUser().getEmail();
var studentEmailCol = 9;
var yearGroupCol = 15;
var displayNameCol = 1;

var SCRIPT_PROPERTY_KEYS = {
  deletedVideosFolderId: 'DELETED_VIDEOS_FOLDER_ID',
  videoSubmissionFormUrl: 'VIDEO_SUBMISSION_FORM_URL'
};


function getScriptProperty_(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || '';
}


function getRequiredScriptProperty_(key) {
  var value = getScriptProperty_(key);
  if (!value) {
    throw new Error('Missing required Script Property: ' + key);
  }
  return value;
}


function normalizeEmail_(email) {
  return email === null || email === undefined ? '' : String(email).trim().toLowerCase();
}


function parseDriveVideo_(value) {
  var rawValue = value === null || value === undefined ? '' : String(value).trim();
  if (!rawValue || rawValue === 'Not Posted Yet') {
    return {
      rawUrl: rawValue,
      fileId: '',
      previewUrl: '',
      directUrl: '',
      openUrl: '',
      resourceKey: '',
      isValid: false,
      status: rawValue === 'Not Posted Yet' ? 'not_posted' : 'missing'
    };
  }

  var fileId = '';
  var resourceKey = '';

  var directMatch = rawValue.match(/\/file\/d\/([a-zA-Z0-9_-]{10,})/);
  if (directMatch && directMatch[1]) {
    fileId = directMatch[1];
  }

  if (!fileId) {
    var idQueryMatch = rawValue.match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
    if (idQueryMatch && idQueryMatch[1]) {
      fileId = idQueryMatch[1];
    }
  }

  if (!fileId) {
    var genericMatch = rawValue.match(/([a-zA-Z0-9_-]{25,})/);
    if (genericMatch && genericMatch[1]) {
      fileId = genericMatch[1];
    }
  }

  var resourceKeyMatch = rawValue.match(/[?&]resourcekey=([^&#]+)/i);
  if (resourceKeyMatch && resourceKeyMatch[1]) {
    resourceKey = resourceKeyMatch[1];
  }

  if (!fileId) {
    return {
      rawUrl: rawValue,
      fileId: '',
      previewUrl: '',
      directUrl: '',
      openUrl: rawValue,
      resourceKey: resourceKey,
      isValid: false,
      status: 'invalid'
    };
  }

  var encodedResourceKey = resourceKey ? '&resourcekey=' + encodeURIComponent(resourceKey) : '';
  var previewResourceKey = resourceKey ? '?resourcekey=' + encodeURIComponent(resourceKey) : '';

  return {
    rawUrl: rawValue,
    fileId: fileId,
    previewUrl: 'https://drive.google.com/file/d/' + fileId + '/preview' + previewResourceKey,
    directUrl: 'https://drive.google.com/uc?export=download&id=' + fileId + encodedResourceKey,
    openUrl: 'https://drive.google.com/file/d/' + fileId + '/view' + previewResourceKey,
    resourceKey: resourceKey,
    isValid: true,
    status: 'ready'
  };
}


// Global variables
var ss = SpreadsheetApp.getActiveSpreadsheet();
var cloudResponseSheet = ss.getSheetByName("Cloud Response");
var studentsSheet = ss.getSheetByName("STUDENTS");
var approverSheet = ss.getSheetByName("Teacher");


var sdata = []; //student data
var cdata = []; //cloud response data

function doGet() {
  var template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('Video Submission Viewer')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


function getUserData() {
  var userData = {
    email: thisUser ,
    isStudent: false,
    isParent: false,
    isApprover: false,
    displayName: '',
    yearGroup: '',
    video: '',
    videoSubmissionFormUrl: getScriptProperty_(SCRIPT_PROPERTY_KEYS.videoSubmissionFormUrl),
    childrenVideos: [],
    allVideos: []
  };
  
  // Check if user is an approver
  if (isApprover(thisUser)) {
    sdata = studentsSheet.getDataRange().getValues();
    cdata = cloudResponseSheet.getDataRange().getValues();
    userData.isApprover = true;
    userData.allVideos = getAllVideos();
  } else {
    // Check if user is a student
    sdata = studentsSheet.getDataRange().getValues();
    var studentData = findStudentByEmail(thisUser );
    if (studentData) {
      cdata = cloudResponseSheet.getDataRange().getValues();
      userData.isStudent = true;
      userData.displayName = studentData.displayName;
      userData.yearGroup = studentData.yearGroup;
      userData.video = parseDriveVideo_(findVideoByEmail(thisUser ));
    } else {
      // Check if user is a parent
      var parentData = findParentByEmail(thisUser );
      if (parentData) {
        cdata = cloudResponseSheet.getDataRange().getValues();
        userData.isParent = true;
        userData.displayName = thisUser;
        userData.childrenVideos = getChildrenVideos(thisUser );
      }
    } 
  }
  Logger.log({
    email: userData.email,
    isStudent: userData.isStudent,
    isParent: userData.isParent,
    isApprover: userData.isApprover,
    yearGroup: userData.yearGroup,
    childrenCount: userData.childrenVideos.length,
    allVideosCount: userData.allVideos.length
  });

  return userData;
}

function isApprover(email) {
  var normalizedEmail = normalizeEmail_(email);
  var data = approverSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalizeEmail_(data[i][0]) === normalizedEmail) {
      return true;
    }
  }
  return false;
}

function getAllVideos() {
  var videos = [];
  for (var i = 1; i < sdata.length; i++) {
    var studentEmail = sdata[i][10];
    var studentData = {
      displayName: sdata[i][2],
      email: studentEmail,
      homeroom: sdata[i][1]
    };
    var video = findVideoByEmail(studentEmail);
    if (video) {
      studentData.video = parseDriveVideo_(video);
    } else {
      studentData.video = parseDriveVideo_('Not Posted Yet');
    }
    videos.push(studentData);
  }
  return videos;
}

function findStudentByEmail(email) {
  var normalizedEmail = normalizeEmail_(email);
  //var data = studentsSheet.getDataRange().getValues();
  var data = sdata;
  for (var i = 1; i < data.length; i++) {
    if (normalizeEmail_(data[i][10]) === normalizedEmail) {
      return {
        displayName: data[i][2],
        yearGroup: data[i][16]
      };
    }
  }
  return null;
}

function findParentByEmail(email) {
  var normalizedEmail = normalizeEmail_(email);
  //var data = studentsSheet.getDataRange().getValues();
  var data = sdata;
  for (var i = 1; i < data.length; i++) {
    if (normalizeEmail_(data[i][9]) === normalizedEmail) {
      return true;
    }
  }
  return false;
}

function findVideoByEmail(email) {
  var normalizedEmail = normalizeEmail_(email);
  //var data = cloudResponseSheet.getDataRange().getValues();
  var data = cdata; 
  for (var i = 1; i < data.length; i++) {
    if (normalizeEmail_(data[i][1]) === normalizedEmail) {
      return data[i][2];
    }
  }
  return '';
} 

function getChildrenVideos(parentEmail) {
  var normalizedParentEmail = normalizeEmail_(parentEmail);
  var videos = [];
  var studentsData = sdata;

  for (var i = 1; i < studentsData.length; i++) {
    if (normalizeEmail_(studentsData[i][9]) === normalizedParentEmail) {
      var studentEmail = studentsData[i][10];
      var video = findVideoByEmail(studentEmail);
      if (video) {
        videos.push({
          displayName: studentsData[i][2],
          video: parseDriveVideo_(video)
        });
      }
    }
  }
  return videos;
}

// -----
// include - include files
// -----
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}


// -----
// Process duplicate Responses when form is submitted.
// -----
function processDuplicateResponses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var responseSheet = ss.getSheetByName("Cloud Response");
  var deletedSheet = ss.getSheetByName("DeletedVideos");
  
  // Get all data from the response sheet
  var data = responseSheet.getDataRange().getValues();
  var headers = data.shift(); // Remove and store headers
  
  // Create an object to store the latest entry for each email
  var latestEntries = {};
  var duplicates = [];
  
  // Process the data to find duplicates and keep the latest entry
  data.forEach(function(row, index) {
    var timestamp = new Date(row[0]);
    var email = row[1];
    var fileUrl = row[2];
    
    if (email in latestEntries) {
      if (timestamp > latestEntries[email].timestamp) {
        duplicates.push(latestEntries[email]);
        latestEntries[email] = {row: row, index: index, timestamp: timestamp};
      } else {
        duplicates.push({row: row, index: index, timestamp: timestamp});
      }
    } else {
      latestEntries[email] = {row: row, index: index, timestamp: timestamp};
    }
  });
  
  // Sort duplicates by index in descending order to avoid shifting rows when deleting
  duplicates.sort(function(a, b) {
    return b.index - a.index;
  });
  
  // Process duplicates
  duplicates.forEach(function(duplicate) {
    // Delete the row from the response sheet
    responseSheet.deleteRow(duplicate.index + 2); // +2 to account for headers and 1-based index
    
    // Append the row to the deleted sheet
    deletedSheet.appendRow(duplicate.row);
    
    // Move the file to the deleted videos folder
    moveFileToDeletionFolder(duplicate.row[2]);
  });
}

function moveFileToDeletionFolder(fileUrl) {
  try {
    var fileId = fileUrl.match(/[-\w]{25,}/);
    if (fileId) {
      var file = DriveApp.getFileById(fileId[0]);
      var folder = DriveApp.getFolderById(getRequiredScriptProperty_(SCRIPT_PROPERTY_KEYS.deletedVideosFolderId));
      file.moveTo(folder);
    }
  } catch (e) {
    Logger.log('Error moving file: ' + e.toString());
  }
}

function batchMoveFilesFromSourceToDestination() {
  var sourceFolderId = '18AtkdUm8_-Nzd66t-L-OPIvOPDpSoLVMfTi_HlAQatP4SXiZqybywAoyiZqHPm_bzBRYh4_o';
  var destinationFolderId = '1adpx8QqNBQrpyxxM7fcZHT-rTwpGB2TZ';

  var sourceFolder = DriveApp.getFolderById(sourceFolderId);
  var destinationFolder = DriveApp.getFolderById(destinationFolderId);

  var files = [];
  var fileIterator = sourceFolder.getFiles();

  while (fileIterator.hasNext()) {
    files.push(fileIterator.next());
  }

  var movedCount = 0;
  for (var i = 0; i < files.length; i++) {
    try {
      files[i].moveTo(destinationFolder);
      movedCount++;
      Logger.log('Moved file: ' + files[i].getName());
    } catch (e) {
      Logger.log('Error moving file: ' + files[i].getName() + ' - ' + e.toString());
    }
  }

  Logger.log('Moved ' + movedCount + ' file(s) from ' + sourceFolder.getName() + ' to ' + destinationFolder.getName());

  return {
    sourceFolderName: sourceFolder.getName(),
    destinationFolderName: destinationFolder.getName(),
    filesFound: files.length,
    movedCount: movedCount
  };
}

function onFormSubmit(e) {
  // Trigger the duplicate processing function when a form is submitted
  processDuplicateResponses();
}