var ss = SpreadsheetApp.getActiveSpreadsheet();

var sheetDetailsList = [
  ['1.List Org Units','Existing organizational Units',],
  ['2.Create Org Units','New Organisation Unit Name*','Creation Status'],
  ['3.Delete Org Units','Organisation Unit Name*', 'Deletion Status'],
  ['4.Create Users','First Name*','Last Name','New Primary Email*','Password*','Org Unit*','Employee ID/Admisiion Number','Creation Status'],
  ['5.Update Users','Primary Email*','First Name','Last Name','Org Unit','Employee ID/Admisiion Number','Update Status'],
  ['6.Suspend Users','Email*','Suspension Status'],
  ['7.Suspended Users List','Name', 'Email'],
  ['8.Activate Suspended Users','Email*','Activation Status'],
  ['9.Delete Users','Email*','Deletion Status'],
  ['10.Un Delete Users','Email*','Org Unit','Restore Status'],
  ['11.List All Users','First Name','Last Name','Primary Email','Org Unit','Domain','2SV Enforced','2SV Enrolled','Creation Date','User Unique ID','Status'],
  ['12.Create Aliases','S No','Category','Name','Primary Email*','Create Alias Status','Alias 1*','Alias 2*','Alias 3*','Add more Alias in next columns'],
  ['13.Delete User Aliases','Name','Primary Email*','Alias Deletion Status','Alias 1*','Alias 2*','Alias 3*','Add more Alias in next columns'],
  ['14.List All Users\' Aliases','Provided Email*','Given Email','Number of Aliases','Aliases',],
  ['15.Create Groups','Group Name*','Group Email*','Group Description','Group Creation Status'],
  ['16.Restrict Group Settings', 'Group Email*', 'Restriction Status'],
  ['17.Delete Groups','Group Email*','Delete Status'],
  ['18.List All Groups','Group Name','Group Email','Members Count','Group ID','Description'],
  ['19.Add Members to Groups','S No',	'Category',	'Member Role* [OWNER / MEMBER / MANAGER]',	'Member Email*',	'Member Addition Status',	'Group Email1*',	'Group Email2*',	'Group Email3*',	'Group Email4*',	'Group Email5*',	'Group Email6*',	'Group Email7*',	'Group Email8*',	'Group Email9*','Group Email9*',	'Add More Group Emails in next columns' ],
  ['20.List All Group Members','Group Name','Group Email','Group ID','Total Members','Member Email','Member Role','Member Type'],
  ['21.Remove Group Member','Name','Member Email*','Removal Status','Group Email1*',	'Group Email2*',	'Group Email3*',	'Group Email4*',	'Group Email5*',	'Group Email6*',	'Group Email7*',	'Group Email8*',	'Group Email9*','Group Email9*',	'Add More Group Emails in next columns' ],
  ['22.Create Classroom Courses','Class/Course/Subject Name*','Owner/Admin/Class Teacher Email*','Description Heading','Description','Section','Room','Class Status','Enrollment Code','Course ID'],
  ['23.List All Courses','Course Name','Course ID','Course Description Heading','Course Status','Course Description','Section','Room','Enrollment Code','Course Link','Guardians Enabled?'],
  ['24.Archive Courses','Course Name*','Course Archival Status'],
  ['25.Activate Archived Courses','Course Name*','Course Archival Status'],
  ['26.Delete Courses','Course Name*','Course Deletion Status'],
  ['27.Add Student to Courses','S No','Class-Sec','Student Name', 'Student Email*','Status',
  'Course/Subject 1','Course/Subject 2','Course/Subject 3','Course/Subject 4','Course/Subject 5','Course/Subject 6','Course/Subject 7','Course/Subject 8',
  'Course/Subject 9','Course/Subject 10','Add more course/Subject Names in Next Columns'],
  ['28.Add Teacher to Courses','S No','Role','Teacher Name','Teacher Email*','Status','Course/Subject 1','Course/Subject 2','Course/Subject 3','Course/Subject 4','Course/Subject 5',
  'Course/Subject6','Course/Subject 7','Course/Subject 8','Add more course/Subject Names in Next Columns'],
  ['29.List All Students by Courses','Course Name*','Name','Course ID','Student Name','Student Email',],
  ['30.List All Teachers by courses','Course Name*','Name','Course ID','Teacher Name','Teacher Email',],
  ['31.List All Courses by Students','Student Name','Student Email*','Course Names'],
  ['32.List All Courses by Teachers','Teacher Name','Teacher Email*','Course Names',],
  ['33.List Students in All courses','Course ID',	'Course Name',	'Student Name',	'Student Email'],
  ['34.List Teachers in All courses','Course ID',	'Course Name',	'Teacher Name',	'Teacher Email'],
  ['35.Remove Students from Courses','Student Name', 'Student Email*','Status','Course Name 1*','Course Name 2*','Course Name 3*','Course Name 4*','Course Name 5*',
  'Course Name 6*','Course Name 7*','Course Name 8*','Add more course names in next columns'],
  ['36.Remove Teachers from Courses','Teacher Name','Teacher Email*','Status','Course Name 1*','Course Name 2*','Course Name 3*','Course Name 4*','Course Name 5*',
  'Course Name 6*','Course Name 7*','Course Name 8*','Add more course names in next columns'],
]

function onOpen() {
  if(!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  var ui = SpreadsheetApp.getUi()
  ui.createMenu('Admin Actions')
    .addSubMenu(ui.createMenu('Setup Options')
      .addItem('Setup All Sheets', 'setupSheetFn')
      .addItem('Clear All Sheets', 'clearAllSheetFn')
      // .addItem('Clear Current Sheet', 'clearCurrentSheetFn')
      .addItem('Delete All Sheets', 'deleteAllSheetFn'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Manage organizational units')
      .addItem(sheetDetailsList[0][0], 'listOrgUnitsFn')
      .addItem(sheetDetailsList[1][0], 'createOrgUnitsFn')
      .addItem(sheetDetailsList[2][0], 'deleteOrgUnitsFn'))
    .addSubMenu(ui.createMenu('Manage Users (Create ID & Password)')
      .addItem(sheetDetailsList[3][0], 'createUsersFn')
      .addItem(sheetDetailsList[4][0], 'updateUsersFn')
      .addItem(sheetDetailsList[5][0], 'suspendUsersFn')
      .addItem(sheetDetailsList[6][0], 'suspendedUsersListFn')
      .addItem(sheetDetailsList[7][0], 'activateUsersFn')
      .addItem(sheetDetailsList[8][0], 'deleteUsersFn')
      .addItem(sheetDetailsList[9][0], 'unDeleteUsersFn')
      .addItem(sheetDetailsList[10][0], 'listAllUsersFn'))
    .addSubMenu(ui.createMenu('Manage Alias (Secondary Email)')
      .addItem(sheetDetailsList[11][0], 'createAliasesFn')
      .addItem(sheetDetailsList[12][0], 'deleteUserAliasesFn')
      .addItem(sheetDetailsList[13][0], 'listAllUsersAliasesFn'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Manage Groups')
      .addItem(sheetDetailsList[14][0], 'createGroupsFn')
      .addItem(sheetDetailsList[15][0], 'restrictGroupSettingsFn')
      .addItem(sheetDetailsList[16][0], 'deleteGroupsFn')
      .addItem(sheetDetailsList[17][0], 'listAllGroupsFn'))
    .addSubMenu(ui.createMenu('Manage Group Members')
      .addItem(sheetDetailsList[18][0], 'addMembers2GroupsFn')
      .addItem(sheetDetailsList[19][0], 'listAllGroupMembersFn')
      .addItem(sheetDetailsList[20][0], 'removeGroupMemberFn'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Manage Classrooms')
      .addItem(sheetDetailsList[21][0], 'createCoursesFn')
      .addItem(sheetDetailsList[22][0], 'listAllCoursesFn')
      .addItem(sheetDetailsList[23][0], 'archiveCoursesFn')
      .addItem(sheetDetailsList[24][0], 'activateArchivedCoursesFn')
      .addItem(sheetDetailsList[25][0], 'deleteCoursesFn'))
    .addSubMenu(ui.createMenu('Manage Classroom Students & Teachers')
      .addItem(sheetDetailsList[26][0], 'addStudents2CoursesFn')
      .addItem(sheetDetailsList[27][0], 'addTeachers2CoursesFn')
      .addItem(sheetDetailsList[28][0], 'listAllStudentsByCoursesFn')
      .addItem(sheetDetailsList[29][0], 'listAllTeachersByCoursesFn')
      .addItem(sheetDetailsList[30][0], 'listAllCourseByStudentsFn')
      .addItem(sheetDetailsList[31][0], 'listAllCourseByTeachersFn')
      .addItem(sheetDetailsList[32][0], 'listAllStudentsInAllCoursesFn')
      .addItem(sheetDetailsList[33][0], 'listAllTeachersInAllCoursesFn')
      .addItem(sheetDetailsList[34][0], 'removeStudents4mCoursesFn')
      .addItem(sheetDetailsList[35][0], 'removeTeachers4mCoursesFn'))
    .addSeparator()
    .addItem('Help', 'helpBoxFn')
    .addToUi();
}

function rAndD (){

  var group = GroupsApp.getGroupByEmail("studG1@lmsedu.in");
  console.log(group.getEmail() + ':');
  var users = group.getUsers();
  for (var i = 0; i < users.length; i++) {
    var user = users[i];
    console.log(user.getEmail());
  }
}

function getCourseId(courseName) {
  var courses = Classroom.Courses.list().courses;
  if (courses && courses.length > 0) {
    for (course in courses) {
      if (courseName === courses[course].name) {
        return courses[course].id;
      } 
    }
  } 
  else {
    return undefined;
  }
}

function getUserId (userEmail) {
  var email = 'email='+ userEmail;
  var userId = AdminDirectory.Users.list({customer: 'my_customer',query : email}).users[0].id;
  return userId;
}

function setupSheetFn() {
  for (var i = 0; i < sheetDetailsList.length; i++) {
    var reqSheet = ss.getSheetByName(sheetDetailsList[i][0]);
    if (reqSheet != null) continue;
    ss.setActiveSheet(ss.getSheets()[i])
    ss.insertSheet().setName(sheetDetailsList[i][0])
    for (var j = 1; j < sheetDetailsList[i].length; j++) {
      ss.getSheetByName(sheetDetailsList[i][0])
        .getRange(1,j)
        .setValue(sheetDetailsList[i][j])
        .setBackgroundRGB(225, 225, 225)
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
    }
  }
}

function clearAllSheetFn() {
  for (var i = 0; i < sheetDetailsList.length; i++) {
    var sheet = ss.getSheetByName(sheetDetailsList[i][0]);
    if (sheet === null) continue;
    sheet.deleteRows(2,sheet.getLastRow())
  }
}

function deleteAllSheetFn() {
  var totalReqSheets = sheetDetailsList.length;
  for (var i =0; i < totalReqSheets; i++) {
    if (ss.getSheetByName(sheetDetailsList[i][0]) != null) ss.deleteSheet(ss.getSheetByName(sheetDetailsList[i][0]));
  }
}

function listOrgUnitsFn() {
  if(!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  if(!AdminDirectory) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[0][0]);
  sheet.getRange(2,1,sheet.getLastRow(),1).clear()
  try {
    var orgUnits = AdminDirectory.Orgunits.get("my_customer","/").organizationUnits;
    for (var orgUnitNum = 0; orgUnitNum <= orgUnits.length; orgUnitNum++) {
      sheet.getRange(orgUnitNum+2,1).setValue(orgUnits[orgUnitNum].name);
    }
  }
  catch (e) {
  }
}
    
function createOrgUnitsFn() {
  if(!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  if(!AdminDirectory) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[1][0]);
  sheet.getRange(2,2,sheet.getLastRow(),1).clear()
  for (var sheetRow = 2; sheetRow <= sheet.getLastRow(); sheetRow++){
    var name = sheet.getRange(sheetRow,1).getDisplayValue();
    if (name[0] === "/") name = name.slice(1);
    try{
      AdminDirectory.Orgunits.insert({"name": name,"parentOrgUnitPath": "/"},"my_customer");
      sheet.getRange(sheetRow,2).setValue('Added Successfully').setHorizontalAlignment("center");}
    catch (e) { sheet.getRange(sheetRow,2).setValue(e.message.split(':')[1]).setHorizontalAlignment("center"); 
    }
  }
}

function deleteOrgUnitsFn() {
  if(!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  if(!AdminDirectory) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[2][0]);
  sheet.getRange(2,2,sheet.getLastRow(),1).clear()
  for (var sheetRow = 2; sheetRow <= sheet.getLastRow(); sheetRow++){
    var orgName = sheet.getRange(sheetRow,1).getDisplayValue();
    if (orgName[0] === "/") orgName = orgName.slice(1);
    try {
      AdminDirectory.Orgunits.remove('my_customer',orgName) 
      sheet.getRange(sheetRow,2).setValue('Deleted Successfully').setHorizontalAlignment("center");
    }
    catch (e) { 
      sheet.getRange(sheetRow,2).setValue("Error").setHorizontalAlignment("center");
    }
  }
}

function createUsersFn() {
  if(!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  if(!AdminDirectory) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[3][0]);
  sheet.getRange(2,7,sheet.getLastRow(),1).clear()
  for (var sheetRow = 2; sheetRow <= sheet.getLastRow(); sheetRow++)
  {
    var familyName = sheet.getRange(sheetRow,2).getDisplayValue();
    var orgName = sheet.getRange(sheetRow,5).getDisplayValue();
    if (familyName === '') familyName = '~';
    if (orgName[0] != "/") orgName = '/'+orgName;
    var userObj = {
      name: {givenName: sheet.getRange(sheetRow,1).getDisplayValue(), 
            familyName: familyName},
      primaryEmail: sheet.getRange(sheetRow,3).getDisplayValue(),
      password: sheet.getRange(sheetRow,4).getDisplayValue(),
      orgUnitPath: orgName,
      changePasswordAtNextLogin: false,
      externalIds: [{type: "organization", value: sheet.getRange(sheetRow,6).getDisplayValue()}],
    };

    try {
      AdminDirectory.Users.insert(userObj);
      sheet.getRange(sheetRow,7).setValue('Created Successfully').setHorizontalAlignment("center");
    }
    catch (e) {
      Logger.log(e)
      sheet.getRange(sheetRow,7).setValue(e.message.split(":")[1]).setHorizontalAlignment("center");
    }
  }
}

function updateUsersFn() {
  if(!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  if(!AdminDirectory) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[4][0]);
  sheet.getRange(2,6,sheet.getLastRow(),1).clear()
  
  for (var sheetRow = 2; sheetRow <= sheet.getLastRow(); sheetRow++) {
    var userKey = sheet.getRange(sheetRow,1).getDisplayValue();
    var orgName = sheet.getRange(sheetRow,4).getDisplayValue();
    if (orgName[0] != "/") orgName = '/'+orgName;
    var userObj = {
      name: {givenName: sheet.getRange(sheetRow,2).getDisplayValue(), 
            familyName: sheet.getRange(sheetRow,3).getDisplayValue()},
      orgUnitPath: orgName,
      externalIds: [{
        type: "organization", 
        value: sheet.getRange(sheetRow,5).getDisplayValue()
        }],
    };

    try{
      AdminDirectory.Users.update(userObj,userKey);
      sheet.getRange(sheetRow,6).setValue('Updated Successfully').setHorizontalAlignment("center");
    }
    catch (e) {
      Logger.log(e)
      sheet.getRange(sheetRow,6).setValue(e.message.split(":")[1]).setHorizontalAlignment("center");
    }
  }
}

function suspendUsersFn() {
  if(!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  if(!AdminDirectory) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[5][0]);
  sheet.getRange(2,2,sheet.getLastRow(),1).clear()
  for (var sheetRow = 2; sheetRow <= sheet.getLastRow(); sheetRow++){
    try{
      AdminDirectory.Users.update({suspended: true},sheet.getRange(sheetRow,1).getDisplayValue());
      sheet.getRange(sheetRow,2).setValue('Suspended Successfully').setHorizontalAlignment("center");
    }
    catch (e) {
      Logger.log(e)
      sheet.getRange(sheetRow,2).setValue(e.message.split(":")[1]).setHorizontalAlignment("center");
    }
  }
}

function suspendedUsersListFn() {
  if(!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  if(!AdminDirectory) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[6][0]);
  sheet.getRange(2,1,sheet.getLastRow(),2).clear()
  var userObj = {
    customer: 'my_customer',
    maxResults: 500,
    query : 'isSuspended=true',
    }
  try {
    var suspendedUsersList = AdminDirectory.Users.list(userObj).users;
    for (var user = 0; user < suspendedUsersList.length; user++) {
      sheet.getRange(user+2,1).setValue(suspendedUsersList[user].name.givenName);
      sheet.getRange(user+2,2).setValue(suspendedUsersList[user].primaryEmail);
    }
  } catch (e) {}
}

function activateUsersFn() {
  if(!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  if(!AdminDirectory) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[7][0]);
  sheet.getRange(2,2,sheet.getLastRow(),1).clear();
  for (var sheetRow = 2; sheetRow <= sheet.getLastRow(); sheetRow++){
    try{
      AdminDirectory.Users.update({suspended: false},sheet.getRange(sheetRow,1).getDisplayValue());
      sheet.getRange(sheetRow,2).setValue('Reactivated Successfully').setHorizontalAlignment("center");
    }
    catch (e) {
      Logger.log(e)
      sheet.getRange(sheetRow,2).setValue(e.message.split(":")[1]).setHorizontalAlignment("center");
    }
  }
}

function deleteUsersFn() {
  if(!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  if(!AdminDirectory) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[8][0]);
  sheet.getRange(2,2,sheet.getLastRow(),1).clear();
  for (var sheetRow = 2; sheetRow <= sheet.getLastRow(); sheetRow++){
    try {
      AdminDirectory.Users.remove(sheet.getRange(sheetRow,1).getDisplayValue());
      sheet.getRange(sheetRow,2).setValue('Deleted Successfully').setHorizontalAlignment("center");
    }
    catch (e) {
      Logger.log(e)
      sheet.getRange(sheetRow,2).setValue(e.message.split(":")[1]).setHorizontalAlignment("center");
    }
  }
}

function unDeleteUsersFn() {
  if(!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  if(!AdminDirectory) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[9][0]);
  sheet.getRange(2,3,sheet.getLastRow(),1).clear();
  var userObj = {
    customerId: 'my_customer',
    maxResults: 500,
    showDeleted: true,
    }
  var deletedUsers = AdminDirectory.Users.list(userObj).users
  for (var sheetRow = 2; sheetRow <= sheet.getLastRow(); sheetRow++){
    var orgName = sheet.getRange(sheetRow,2).getDisplayValue();
    if (orgName[0] != "/") orgName = '/'+orgName;
    try{
      for (var dltUsr = 0; dltUsr < deletedUsers.length; dltUsr++) {
        if (sheet.getRange(sheetRow,1).getDisplayValue() === deletedUsers[dltUsr].primaryEmail) {
          AdminDirectory.Users.undelete({orgUnitPath: orgName},deletedUsers[dltUsr].id);
          sheet.getRange(sheetRow,3).setValue('Restored Successfully').setHorizontalAlignment("center");
          break
        }
      }
    }
    catch (e) {
      Logger.log(e)
      sheet.getRange(sheetRow,3).setValue(e.message.split(":")[1]).setHorizontalAlignment("center");
    }
  }
}

function listAllUsersFn() {
  if(!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  if(!AdminDirectory) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[10][0]);
  sheet.getRange(2,1,sheet.getLastRow(),sheet.getLastColumn()).clear();
  var pageToken;
  var page;
  var allUsersList = []
  do {
    page = AdminDirectory.Users.list({
      customer: 'my_customer',
      orderBy: 'givenName',
      maxResults: 500,
      pageToken: pageToken
    });
    var users = page.users;
    if (users) {
      for (var i = 0; i < users.length; i++) {
        var user = users[i];
        var userStatus = ''
        if (user.isAdmin == true || user.isDelegatedAdmin == true) {
          userStatus = 'Admin'
        }
        else if (user.suspended == true) {
          userStatus = 'Suspended'
        }
        else if (user.archived == true) {
          userStatus = 'Archived'
        }
        else if (user.showDeleted == true) {
          userStatus = 'Deleted'
        }
        allUsersList.push([
          user.name.givenName, 
          user.name.familyName,
          user.primaryEmail,
          user.orgUnitPath,
          user.primaryEmail.split('@')[1],
          user.isEnforcedIn2Sv,
          user.isEnrolledIn2Sv,
          user.creationTime.slice(0,10),
          user.id,
          userStatus])
      }
    } else {
      Logger.log('No users found.');
    }
    pageToken = page.nextPageToken;
  } while (pageToken);
  sheet.getRange(2,1,allUsersList.length,allUsersList[0].length).setValues(allUsersList)
}

function addInfoPrompt() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var result = ui.alert(
     'Info',
     '1. If a user is newly added, the name will become green \n2. If the user is already added, it will become blue \n3. If there is any error in adding, it will become red color  \n\nDo you want to proceed? Click "OK"',
      ui.ButtonSet.OK);
  if (result == ui.Button.OK) {return 'OK'}
}

function addUserActions(sheet, type){
  if (!SpreadsheetApp) return;
  for (var sheetRow = 2; sheetRow<=sheet.getLastRow(); sheetRow++) {
    var statusCell = sheet.getRange(sheetRow,5).getDisplayValue();
    if (statusCell == 'Added Successfully' || statusCell == 'Already Exists') {continue};
    for (var sheetCol=6; sheetCol<=sheet.getLastColumn(); sheetCol++) {
      if (sheet.getRange(sheetRow,sheetCol).getDisplayValue() == '') {break}
      try {
        if (type == 'Students') {
          Classroom.Courses.Students.create({
            'userId': sheet.getRange(sheetRow,4).getDisplayValue()}, 
            getCourseId(sheet.getRange(sheetRow,sheetCol).getDisplayValue()));
        }
        if (type == 'Teachers') {
          Classroom.Courses.Teachers.create({
            'userId': sheet.getRange(sheetRow,4).getDisplayValue()}, 
            getCourseId(sheet.getRange(sheetRow,sheetCol).getDisplayValue()));
        }
        if (type == 'ALIAS') {
          AdminDirectory.Users.Aliases.insert({
            alias: sheet.getRange(sheetRow,sheetCol).getDisplayValue()}, 
            sheet.getRange(sheetRow,4).getDisplayValue());
        }
        if (type == 'GROUP') {
          AdminDirectory.Members.insert({
            email: sheet.getRange(sheetRow,4).getDisplayValue(), 
            role: sheet.getRange(sheetRow,3).getDisplayValue()},
            sheet.getRange(sheetRow,sheetCol).getDisplayValue())
        }
        sheet.getRange(sheetRow,sheetCol).setFontColor("green").setFontWeight('bold')
        if (sheet.getRange(sheetRow,sheetCol+1).getDisplayValue() == '' && statusCell == '') {
          sheet.getRange(sheetRow,5).setValue('Added Successfully').setHorizontalAlignment("center")};
        }
      catch (e) {Logger.log(e);
        if (
          (e == "GoogleJsonResponseException: API call to classroom.courses.teachers.create failed with error: Requested entity already exists" 
          || e == 'GoogleJsonResponseException: API call to directory.users.aliases.insert failed with error: Entity already exists.'
          || e == 'GoogleJsonResponseException: API call to directory.members.insert failed with error: Member already exists.')
          && sheet.getRange(sheetRow,sheetCol+1).getDisplayValue() == ''
          && statusCell == '') {
            sheet.getRange(sheetRow,5).setValue("Already Exists").setHorizontalAlignment("center");
            sheet.getRange(sheetRow,sheetCol).setFontColor('blue').setFontWeight('bold')
        }
        if (
          e == "GoogleJsonResponseException: API call to classroom.courses.teachers.create failed with error: Requested entity already exists" 
          || e == 'GoogleJsonResponseException: API call to directory.users.aliases.insert failed with error: Entity already exists.'
          || e == 'GoogleJsonResponseException: API call to directory.members.insert failed with error: Member already exists.') {
            sheet.getRange(sheetRow,sheetCol).setFontColor('blue').setFontWeight('bold');
        }
        else {
          sheet.getRange(sheetRow,5).setValue(e.message.split(':')[1]).setHorizontalAlignment("center");
          sheet.getRange(sheetRow,sheetCol).setFontColor('red').setFontWeight('bold')
        }
      }
    }
  }
}

function createAliasesFn() {
  if (!Classroom) return; //Add Classroom API at the Sevices (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[11][0]);
  if(addInfoPrompt() == 'OK'){
    addUserActions(sheet,'ALIAS')
  }
}

function removeInfoPrompt() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var result = ui.alert(
     'Info',
     '1. If it is currently removed, the name will become green \n2. If it is already removed before or the name does not exist, it will become red   \n\nDo you want to proceed? Click "OK"',
      ui.ButtonSet.OK);
  if (result == ui.Button.OK) {return 'OK'}
}

function removeUserActions(sheet, type) {
  for (var sheetRow = 2; sheetRow<=sheet.getLastRow(); sheetRow++) {
    if (sheet.getRange(sheetRow,3).getDisplayValue() == 'Removed Successfully' || sheet.getRange(sheetRow,3).getDisplayValue() == 'Already Removed') {continue};
    for (var sheetCol=4; sheetCol<=sheet.getLastColumn(); sheetCol++) {
      if (sheet.getRange(sheetRow,sheetCol).getDisplayValue() == '') {break}
      try {
        if (type == 'Students') {
          Classroom.Courses.Students.remove(getCourseId(sheet.getRange(sheetRow,sheetCol).getDisplayValue()), sheet.getRange(sheetRow,2).getDisplayValue());
        }
        if (type == 'Teachers') {
          Classroom.Courses.Teachers.remove(getCourseId(sheet.getRange(sheetRow,sheetCol).getDisplayValue()), sheet.getRange(sheetRow,2).getDisplayValue());
        }
        if (type == 'ALIAS') {
          AdminDirectory.Users.Aliases.remove(sheet.getRange(sheetRow,2).getDisplayValue(), sheet.getRange(sheetRow,sheetCol).getDisplayValue());
        }
        if (type == 'MEMBER') {
          AdminDirectory.Members.remove(sheet.getRange(sheetRow,sheetCol).getDisplayValue(),sheet.getRange(sheetRow,2).getDisplayValue())
        }
        sheet.getRange(sheetRow,sheetCol).setFontColor("green").setFontWeight('bold')
        if (sheet.getRange(sheetRow,sheetCol+1).getDisplayValue() == '') {
          sheet.getRange(sheetRow,3).setValue('Removed Successfully');
        }
      }
      catch (e) { Logger.log(e)
        if (e.message.split(':')[1] == ' @UserCannotUnenrollFromCourse This user does not have permission to unenroll from the course.') {
          sheet.getRange(sheetRow,3).setValue(e)
          var ui = SpreadsheetApp.getUi(); // Same variations.
          var result = ui.alert( 'SORRY', 'Your admin has disabled Unenrollment permissions for students and Teachers \n\nGo to https://admin.google.com/ \n\nGo to Apps -> Google Workspace -> Classroom -> Student unenrollment \n\nChange "Who can unenroll students from classes?" from "Teachers only" to "Students and teachers"' , ui.ButtonSet.OK);
          if (result == ui.Button.OK) {break}
        }
        if (e == 'GoogleJsonResponseException: API call to directory.users.aliases.delete failed with error: Invalid Input: resource_id' 
        || e == 'GoogleJsonResponseException: API call to classroom.courses.teachers.delete failed with error: Requested entity was not found.'
        || e == 'GoogleJsonResponseException: API call to directory.members.delete failed with error: Resource Not Found: memberKey') {
          sheet.getRange(sheetRow,3).setValue(e.message.split(':')[1]).setHorizontalAlignment("center");
          sheet.getRange(sheetRow,sheetCol).setFontColor('red').setFontWeight('bold')
        }
        else {
          sheet.getRange(sheetRow,sheetCol).setFontColor('grey').setFontWeight('bold');
          if (sheet.getRange(sheetRow,sheetCol+1).getDisplayValue() == '')
            sheet.getRange(sheetRow,3).setValue('Error');
        }
      }
    }
  }
}

function deleteUserAliasesFn() {
  if (!AdminDirectory) return; //Add Classroom API at the Sevices (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[12][0]);
  if(removeInfoPrompt() == 'OK'){
    removeUserActions(sheet,'ALIAS')
  }
}

function listAllUsersAliasesFn () {
  var sheet = ss.getSheetByName(sheetDetailsList[13][0]);
  sheet.getRange(2,2,sheet.getLastRow(),3).clear();
  var i = 2
  for (var sheetRow=2; sheetRow<=sheet.getLastRow(); sheetRow++) {
    if(sheet.getRange(sheetRow,1).getDisplayValue() == '') break;
    var aliasesList = AdminDirectory.Users.Aliases.list(sheet.getRange(sheetRow,1).getDisplayValue()).aliases;
    sheet.getRange(i,2).setValue(sheet.getRange(sheetRow,1).getDisplayValue())
    try {
      aliasesList.length
    } 
    catch (e) {
      Logger.log (e)
      sheet.getRange(i,3).setValue(0)
      i += 1
      continue
    }
    sheet.getRange(i,3).setValue(aliasesList.length)
    for (alias in aliasesList) {
      sheet.getRange(i,4).setValue(aliasesList[alias].alias)
      i += 1
    }
  }
}

function createGroupsFn () {
  if(!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  if(!AdminDirectory) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[14][0]);
  sheet.getRange(2,4,sheet.getLastRow(),1).clear();
  for (var sheetRow=2; sheetRow<=sheet.getLastRow(); sheetRow++) {
    try {
      AdminDirectory.Groups.insert({
        name: sheet.getRange(sheetRow,1).getDisplayValue(), 
        email: sheet.getRange(sheetRow,2).getDisplayValue(), 
        description: sheet.getRange(sheetRow,3).getDisplayValue()})
      sheet.getRange(sheetRow,4).setValue("Group Created").setHorizontalAlignment("center");
    } catch (e) {
      Logger.log(e)
      sheet.getRange(sheetRow,4).setValue(e.message.split(':')[1])
    }
  } 
}

function restrictGroupSettingsFn() {
  var sheet = ss.getSheetByName(sheetDetailsList[15][0]);
  sheet.getRange(2,2,sheet.getLastRow(),1).clear();
  for (var sheetRow=2; sheetRow<=sheet.getLastRow(); sheetRow++) {
    try {
      var groupId = sheet.getRange(sheetRow,1).getDisplayValue();
      var groupObj = {
        whoCanContactOwner : 'ALL_OWNERS_CAN_CONTACT',
      whoCanViewMembership : 'ALL_IN_DOMAIN_CAN_VIEW',
      whoCanViewGroup : 'ALL_OWNERS_CAN_VIEW',
      whoCanPostMessage : 'NONE_CAN_POST',
      whoCanApproveMembers : 'NONE_CAN_APPROVE',
      whoCanJoin : 'INVITED_CAN_JOIN',
      allowExternalMembers : false,
      whoCanLeaveGroup : 'NONE_CAN_LEAVE',
      whoCanAdd : 'NONE_CAN_ADD',
      whoCanModerateMembers : 'NONE',
      whoCanDiscoverGroup : 'ALL_IN_DOMAIN_CAN_DISCOVER',
      whoCanPostAnnouncements : 'OWNERS_ONLY',
      archiveOnly : true,
      }
      AdminGroupsSettings.Groups.update(groupObj,groupId)
    } catch (e) {Logger.log (e)}
  }
}

function deleteGroupsFn () {
  if(!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  if(!AdminDirectory) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[16][0]);
  sheet.getRange(2,2,sheet.getLastRow(),1).clear();
  for (var sheetRow=2; sheetRow<=sheet.getLastRow(); sheetRow++) {
    try {
      AdminDirectory.Groups.remove(sheet.getRange(sheetRow,1).getDisplayValue());
      sheet.getRange(sheetRow,2).setValue("Group Deleted").setHorizontalAlignment("center");
    } catch (e) {
      Logger.log(e)
      sheet.getRange(sheetRow,2).setValue(e.message.split(':')[1])
    }
  } 
}

function listAllGroupsFn(sendResultToSheet = false) {
  if(!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  if(!AdminDirectory) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[17][0]);
  sheet.getRange(2,1,sheet.getLastRow(),sheet.getLastColumn()).clear();
  var allGroupsList = []
  var pageToken;
  var page;
  do {
    page = AdminDirectory.Groups.list({
      customer: 'my_customer',
      maxResults: 500,
      pageToken: pageToken
    });
    var groups = page.groups;
    if (groups) {
      for (var i = 0; i < groups.length; i++) {
        var group = groups[i];
        allGroupsList.push([
          group.name,
          group.email,
          group.directMembersCount,
          group.id,
          group.description,
        ])
      }
    } else {
      Logger.log('No groups found.');
    }
    pageToken = page.nextPageToken;
  } while (pageToken);
  if (!sendResultToSheet) sheet.getRange(2,1,allGroupsList.length,allGroupsList[0].length).setValues(allGroupsList);
  else return allGroupsList;
}

function addMembers2GroupsFn() {
  if (!AdminDirectory) return; //Add Classroom API at the Sevices (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[18][0]);
  if(addInfoPrompt() == 'OK'){
    addUserActions(sheet,'GROUP')
  }
}

function listAllGroupMembersFn() {
  if(!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  if(!AdminDirectory) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[19][0]);
  sheet.getRange(2,1,sheet.getLastRow(),sheet.getLastColumn()).clear();
  var sheetRow = 2
  var groupsList = listAllGroupsFn(true)
  for (group in groupsList) {
    var groupMembersList = []
    sheet.getRange(sheetRow,1).setValue(groupsList[group][0])
    sheet.getRange(sheetRow,2).setValue(groupsList[group][1])
    sheet.getRange(sheetRow,3).setValue(groupsList[group][3])
    var pageToken;
    var page;
    do {
      page = AdminDirectory.Members.list(
        groupsList[group][1],
        {customer: 'my_customer', maxResults: 500, pageToken: pageToken});
      var members = page.members;
      if (members) {
        for (var i = 0; i < members.length; i++) {
          var member = members[i];
          groupMembersList.push([
            member.email,
            member.role,
            member.type
          ])
        }
      } else { Logger.log('No members found.');}
      pageToken = page.nextPageToken;
    } while (pageToken);
    if (groupMembersList.length > 0) {
      sheet.getRange(sheetRow,4).setValue(groupMembersList.length)
      for (member in groupMembersList) {
        sheet.getRange(sheetRow,5).setValue(groupMembersList[member][0]);
        sheet.getRange(sheetRow,6).setValue(groupMembersList[member][1]);
        sheet.getRange(sheetRow,7).setValue(groupMembersList[member][2]);
        sheetRow += 1
      }
    } else {
      sheet.getRange(sheetRow,4).setValue(0);
      sheetRow += 1
    }
  }
}

function removeGroupMemberFn() {
  if (!AdminDirectory) return; //Add Classroom API at the Sevices (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[20][0]);
  if(removeInfoPrompt() == 'OK'){
    removeUserActions(sheet,'MEMBER')
  }
}

function createCoursesFn() {
  if (!Classroom) return; //Add Classroom API at the Sevices +  (!Look at the left pane)
  if(!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)

  var sheet = ss.getSheetByName(sheetDetailsList[21][0]);

  var courses = Classroom.Courses.list().courses;
  sheet.getRange(2,7,sheet.getLastRow(),1).clear();
  for (var sheetRow=2; sheetRow<=sheet.getLastRow(); sheetRow++) { 
    for (course in courses) {
      if (sheet.getRange(sheetRow,1).getDisplayValue() === courses[course].name) {
        sheet.getRange(sheetRow,7).setValue("Already Exist").setHorizontalAlignment("center");
        break
      }
    }
    if (sheet.getRange(sheetRow,7).getDisplayValue() != "Already Exist") {
      var courseObj = {
        name: sheet.getRange(sheetRow,1).getDisplayValue(),
        ownerId: sheet.getRange(sheetRow,2).getDisplayValue(),
        descriptionHeading: sheet.getRange(sheetRow,3).getDisplayValue(),    
        description: sheet.getRange(sheetRow,4).getDisplayValue(),
        section: sheet.getRange(sheetRow,5).getDisplayValue(),    
        room: sheet.getRange(sheetRow,6).getDisplayValue(),
        courseState: 'ACTIVE'
      };
      try {
        var course = Classroom.Courses.create(courseObj);
        sheet.getRange(sheetRow,7).setValue("Course Created").setHorizontalAlignment("center");
        sheet.getRange(sheetRow,8).setValue(course.enrollmentCode).setHorizontalAlignment("center");
        sheet.getRange(sheetRow,9).setValue(course.id).setHorizontalAlignment("center");
      }
      catch (e) {
        sheet.getRange(sheetRow,7).setValue(e.message.split(':')[1]).setHorizontalAlignment("center");
      }
    }
  }
}

function listAllCoursesFn() {
  if (!Classroom) return; //Add Classroom API at the Sevices +  (!Look at the left pane)
  if(!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)

  var sheet = ss.getSheetByName(sheetDetailsList[22][0]);
  sheet.getRange(2,1,sheet.getLastRow(),sheet.getLastColumn()).clear();
  var allCoursesList = [];
  var courses = Classroom.Courses.list().courses;
  for (course in courses) {
    var course = courses[course];
    try{
      allCoursesList.push([
        course.name,
        course.id,
        course.descriptionHeading,
        course.courseState,
        course.description,
        course.section,
        course.room,
        course.enrollmentCode,
        course.alternateLink,
        course.guardiansEnabled]);
    } catch (e) {}
  }
  sheet.getRange(2,1,allCoursesList.length,allCoursesList[0].length).setValues(allCoursesList);
}

function archiveCoursesFn() {
  if (!Classroom) return; //Add Classroom API at the Sevices +  (!Look at the left pane)
  if(!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)

  var sheet = ss.getSheetByName(sheetDetailsList[23][0]);
  sheet.getRange(2,2,sheet.getLastRow(),1).clear();

  var courses = Classroom.Courses.list().courses;
  for (var sheetRow=2; sheetRow<=sheet.getLastRow(); sheetRow++) {
    var courseObj = {
        name: sheet.getRange(sheetRow,1).getDisplayValue(),
        courseState: 'ARCHIVED'
      };
    for (course in courses) {
      if (sheet.getRange(sheetRow,1).getDisplayValue() === courses[course].name) {
        try {
          Classroom.Courses.update(courseObj,courses[course].id)
          sheet.getRange(sheetRow,2).setValue("Course Archived").setHorizontalAlignment("center");
        }
        catch (e) {
          sheet.getRange(sheetRow,2).setValue(e.message.split(':')[1]).setHorizontalAlignment("center");
        }
      }
    }
  }
}

function activateArchivedCoursesFn() {
  if (!Classroom) return; //Add Classroom API at the Sevices +  (!Look at the left pane)
  if(!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)

  var sheet = ss.getSheetByName(sheetDetailsList[24][0]);
  sheet.getRange(2,2,sheet.getLastRow(),1).clear();

  var courses = Classroom.Courses.list().courses;
  for (var sheetRow=2; sheetRow<=sheet.getLastRow(); sheetRow++) {
    var courseObj = {
        name: sheet.getRange(sheetRow,1).getDisplayValue(),
        courseState: 'ACTIVE'
      };
    for (course in courses) {
      if (sheet.getRange(sheetRow,1).getDisplayValue() === courses[course].name) {
        try {
          Classroom.Courses.update(courseObj,courses[course].id)
          sheet.getRange(sheetRow,2).setValue("Course Activated").setHorizontalAlignment("center");
        }
        catch (e) {
          sheet.getRange(sheetRow,2).setValue(e.message.split(':')[1]).setHorizontalAlignment("center");
        }
      }
    }
  }
}

function deleteCoursesFn() {
  if (!Classroom) return; //Add Classroom API at the Sevices +  (!Look at the left pane)
  if(!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)

  var sheet = ss.getSheetByName(sheetDetailsList[25][0]);
  sheet.getRange(2,2,sheet.getLastRow(),1).clear();

  for (var sheetRow=2; sheetRow<=sheet.getLastRow(); sheetRow++) {
    try {
      Classroom.Courses.remove(getCourseId(sheet.getRange(sheetRow,1).getDisplayValue()));
      sheet.getRange(sheetRow,2).setValue("Course Deleted").setHorizontalAlignment("center");
    }
    catch (e) {
      sheet.getRange(sheetRow,2).setValue(e.message.split(':')[1]).setHorizontalAlignment("center");
    }
  }
}

function addInfoPrompt() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var result = ui.alert(
     'Info',
     '1. If it is newly added, it will become green \n2. If it is already added, it will become blue \n3. If there is any error in adding, it will become red color  \n\nDo you want to proceed? Click "OK"',
      ui.ButtonSet.OK);
  if (result == ui.Button.OK) {return 'OK'}
}

function addStudents2CoursesFn() {
  if (!Classroom) return; //Add Classroom API at the Sevices (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[26][0]);
  if(addInfoPrompt() == 'OK'){
    addUserActions(sheet,'Students')
  }
}

function addTeachers2CoursesFn() {
  if (!Classroom) return; //Add Classroom API at the Sevices (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[27][0]);
  if(addInfoPrompt() == 'OK'){
    addUserActions(sheet,'Teachers')
  }
}

function allUsersByCourse (sheet, usr) {
  if(!SpreadsheetApp) return;
  var courseNames = []
  for (var sheetRow = 2; sheetRow <= sheet.getLastRow(); sheetRow++) {
    courseNames.push(sheet.getRange(sheetRow,1).getDisplayValue());
  }
  sheet.getRange(2,2,sheet.getLastRow(),sheet.getLastColumn()).clear()
  var allUsersByCoursesList = []
  for (courseName in courseNames) { 
    var courseName = courseNames[courseName]
    var courseId = getCourseId(courseName)
    var nextPageToken = '';
    do {
      var optionalArgs = {pageToken: nextPageToken};
      try {
        if (usr == 'Students') {
          var response = Classroom.Courses.Students.list(courseId, optionalArgs);
          var usersList = response.students;
        }
        if (usr == 'Teachers') {
          var response = Classroom.Courses.Teachers.list(courseId, optionalArgs);
          var usersList = response.teachers;
        }
        nextPageToken = response.nextPageToken;
        for (user in usersList) {
          try {
            allUsersByCoursesList.push([
              courseName,
              courseId,
              usersList[user].profile.name.givenName,
              usersList[user].profile.emailAddress]);
          } catch (e) {}
        } 
      } catch (e) {}
    } while (nextPageToken);
  }
  sheet.getRange(i,2,allUsersByCoursesList.length,allUsersByCoursesList[0].length).setValues(allUsersByCoursesList);
}

function listAllStudentsByCoursesFn() {
  if (!Classroom) return; //Add Classroom API at the Sevices (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[28][0]);
  allUsersByCourse (sheet, 'Students')
}

function listAllTeachersByCoursesFn() {
  if (!Classroom) return; //Add Classroom API at the Sevices (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[29][0]);
  allUsersByCourse (sheet, 'Teachers')
}

function listAllUsersInAllCourses(sheet, usr, sendResultToSheet = true) {
  var courses = Classroom.Courses.list().courses;
  
  var allUsersInAllCoursesList = []
  for (course in courses) {
    var courseId = courses[course].id
    var nextPageToken = '';
    do {
      var optionalArgs = {pageToken: nextPageToken};
      try {
        if (usr == 'Students') {
          var response = Classroom.Courses.Students.list(courseId, optionalArgs);
          var usersList = response.students;
        }
        if (usr == 'Teachers') {
          var response = Classroom.Courses.Teachers.list(courseId, optionalArgs);
          var usersList = response.teachers;
        }
        nextPageToken = response.nextPageToken;
        for (user in usersList) {
          try {
            allUsersInAllCoursesList.push([courseId,courses[course].name,usersList[user].profile.name.givenName,usersList[user].profile.emailAddress])
          } catch (e) {Logger.log(e)}
        }
      } catch (e) {Logger.log(e)}
    } while (nextPageToken);
  }
  if (sendResultToSheet) sheet.getRange(2,1,allUsersInAllCoursesList.length,4).setValues(allUsersInAllCoursesList);
  else return allUsersInAllCoursesList;
}

function listAllCourseByUsers(sheet, usr) {
  if (!Classroom) return; //Add Classroom API at the Sevices +  (!Look at the left pane)
  if (!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  var listOfCoursesAndUsers = listAllUsersInAllCourses(sheet ,usr, false)
  sheet.getRange(2,3,sheet.getLastRow(),sheet.getLastColumn()).clear()
  for (var sheetRow = 2; sheetRow <= sheet.getLastRow(); sheetRow++) {
    var i = 3
    for (list in listOfCoursesAndUsers) {
      if(sheet.getRange(sheetRow,2).getDisplayValue() == listOfCoursesAndUsers[list][3]) {
        sheet.getRange(sheetRow,i).setValue(listOfCoursesAndUsers[list][1]);
        i += 1
      }
    }
  }
}

function listAllCourseByStudentsFn () {
  if (!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)

  var sheet = ss.getSheetByName(sheetDetailsList[30][0]);

  listAllCourseByUsers(sheet, "Students")
}

function listAllCourseByTeachersFn () {
  if (!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)

  var sheet = ss.getSheetByName(sheetDetailsList[31][0]);
  
  listAllCourseByUsers(sheet, "Teachers")
}

function listAllStudentsInAllCoursesFn () {
  if (!Classroom) return; //Add Classroom API at the Sevices +  (!Look at the left pane)
  if (!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[32][0]);
  sheet.getRange(2,1,sheet.getLastRow(),sheet.getLastColumn()).clear()
  listAllUsersInAllCourses(sheet, 'Students')
}

function listAllTeachersInAllCoursesFn () {
  if (!Classroom) return; //Add Classroom API at the Sevices +  (!Look at the left pane)
  if (!SpreadsheetApp) return; //Add Google Sheet API at the Sevices +  (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[33][0]);
  sheet.getRange(2,1,sheet.getLastRow(),sheet.getLastColumn()).clear()
  listAllUsersInAllCourses(sheet, 'Teachers')
}

function removeStudents4mCoursesFn() {
  if (!Classroom) return; //Add Classroom API at the Sevices (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[34][0]);
  if(removeInfoPrompt() == 'OK'){
    removeUserActions(sheet,'Students')
  }
}

function removeTeachers4mCoursesFn() {
  if (!Classroom) return; //Add Classroom API at the Sevices (!Look at the left pane)
  var sheet = ss.getSheetByName(sheetDetailsList[35][0]);
  if(removeInfoPrompt() == 'OK'){
    removeUserActions(sheet,'Teachers')
  }
}

function helpBoxFn() {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var result = ui.alert(
     'HELP',
     '1. Click the "Setup Sheet" from "Setup Options" Menu. \n2. Then use other options\n\nDo you need more help? Click "Yes" or "No"',
      ui.ButtonSet.YES_NO);
  if (result == ui.Button.YES) ui.alert('Send Mail to ganesaperumal@live.com'); 
}
