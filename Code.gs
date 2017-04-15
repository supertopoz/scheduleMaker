
//------------------ Unit Test -------------------------------------

function assert(actual, expected, testName){
    if (actual === expected) Logger.log("Passed");
    else Logger.log("FAILED ["+testName+"] expected: "+expected+" but got: "+ actual)
}


// ----------------- Init actions -----------------------------------

function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu().addItem('Open Check Tool',
    'showSidebar').addToUi();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  var thisDoc = SpreadsheetApp.getActive();
  //   var classCode = thisDoc.getName().substring(0,12);
  var ui = HtmlService.createTemplateFromFile('Sidebar').evaluate().setTitle(
    'Schedule Scanner').setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(ui);
}







// ----------------- Common Tools -------------------------------------------------

function ValidURL(str) {
  var regex = /(http|https):\/\/(\w+:{0,1}\w*)?(\S+)(:[0-9]+)?(\/|\/([\w#!:.?+=&%!\-\/]))?/;
  if(!regex .test(str)) {
//    alert("Please enter valid URL.");
    return false;
  } else {
    return true;
  }
}

function arrayMaker(source) {
    var result = source[0].map(function(elem, i, arr) {
        return [elem, (i + 1 < arr.length) ? arr[i + 1] : null];
    }).filter(function(elem, i) {
        return !(i % 2);
    });
    return result
}


function deleteRow(sheetName){
// var sheetName ="Standards Checker";
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheetByName(sheetName);
 var LastRow = sheet.getMaxRows()-1;
 var all = sheet.getDataRange();
   all.setBackground("white")
   all.setFontColor("black") 
 Logger.log(LastRow);
 // Rows start at "1" - this will delete the first two rows
 sheet.deleteRows(2, LastRow);
 }
 
 
 function getTabs() {
    var sheetName = [];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    for(var item in sheets){
    sheetName.push(sheets[item].getName());
    }
    return sheetName;
}
 
function addMissingTabs(){
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var tabs = getTabs(); 
   if (tabs.indexOf("Standards Checker") <0 ){
    ss.insertSheet("Standards Checker");  
   } 
   if (tabs.indexOf("Assessment") <0 ){
    ss.insertSheet("Assessment");  
   } 
   if (tabs.indexOf("Admin Tasks") <0 ){
    ss.insertSheet("Admin Tasks");  
   }
   
} 

function getAssessmentTabs() {
    var sheetName = [];
    var SPREADSHEET_ID = "1PIMtKuiA9hzxb-9GT0dE7Wvn-4pZC8nrOx_943dk8YQ"
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    //var sheet = ss.getSheetByName("Standards Checker")
  //  var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    for(var item in sheets){
    sheetName.push(sheets[item].getName());
    }
    Logger.log(sheetName)
    return sheetName;
}
 
function stripes(sheet) {
	var count = 1
	var programmes = {};
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var sheet = ss.getSheetByName(sheet);
	var maxCol = sheet.getMaxColumns();
	var all = sheet.getDataRange();
	all.setBackground("white")
	all.setFontColor("black")
	var lrow = sheet.getLastRow();
	var data = sheet.getRange(1, 1, lrow).getDisplayValues()
	for (var i = 0; i < data.length; i++) {
		programmes[data[i]] = i
		Logger.log(Math.floor(i));
	}
	for (var items in programmes) {
		var rowNumber = programmes[items];
		var row = sheet.getRange(rowNumber + 2, 1, 1, maxCol);
		row.setBackground("blue")
		row.setFontColor("white")
	}
}
 
 
// Not yet in App
function fillemptyCells (){
  var column = 1;
  var SPREADSHEET_ID = "1eIBkJB-nlexe91vReHH9QIShnCpkLFafbSpvgRFdjPk"
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("Standards Checker")
  //Get the first cell in the column
  var lastrow = sheet.getLastRow()
  for (var i =2; i<lastrow;i++){
   var cell = sheet.getRange(i,column).getValues(); 
   Logger.log(cell)
   if (cell.toString() === ""){
   //Get the value of the pervious cell. 
    var previousCell = sheet.getRange(i-1,column).getValues(); 
   // Put it into the current cell. 
    var currentCell = sheet.getRange(i,column).setValue(previousCell); 
   }
   }
}

// ----------------- Common Tools finished -------------------------------------------------






// ----------------- Create column of hrs dones -----------------------------------

var totalNoLessons, lessonPWeek;
var text = [];

function addLessons(level){
    var hrsDone = [];
 //   var level = "S1B";
    var ss = SpreadsheetApp.getActiveSpreadsheet();
	var sheet = ss.getSheetByName(level);
    Logger.log("Working");
    Logger.log(level);
    var lastRow = sheet.getLastRow();
   
    var data = sheet.getRange(2,2,lastRow,1).getValues();
    data.forEach(function(items){
      hrsDone.push(items[0])
    })
    return hrsDone
 
//return an array of hrs done

}

// ----------------- Create column of hrs done finished -----------------------------------







// ----------------- Make new schedule -----------------------------------
function makeNewSchedule(level) {
//    var level = ['S2A','S3A','S3B']
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var ssId = ss.getId();
	var sourceHyperlinks = ssId;
	//Create a new schedule and Give it a name 
	var folderIDs = ['0B5miykSsL-gYOVRTUWFrWnE2dWM'];
    //https://drive.google.com/drive/folders/0B5miykSsL-gYOVRTUWFrWnE2dWM
	for (var items in level) {
		var newSchedule = SpreadsheetApp.create(level[items]);
		var scheduleId = newSchedule.getId();
		var schedule = DriveApp.getFileById(scheduleId);
		DriveApp.getFolderById(folderIDs).addFile(schedule);
		DriveApp.getRootFolder().removeFile(schedule);
		// get data
		
        
       //var hrsDone = addLessons(lessonLength, courseLength); 
       var hrsDone = addLessons(level[items])
              
       var data = getData(level[items], sourceHyperlinks);
			//Fill it with correct content. 
       var ss = SpreadsheetApp.openById(scheduleId);
       var sheet = ss.getSheetByName('Sheet1');
       var lessonNo = 1;
		//    var hrsDone = 1.5;
        // Schedule header
		sheet.appendRow(['Class No.', 'Hrs done', 'Date', 'Admin Task', 'Course book page/unit', 'Actual page done/ Supplementary Materials', 'Lesson Aims and/or target language', 'Practice Activity', 'Production Activity', 'Teacher'])
		for (var item in data[0]) {
			var admin = data[1][item].join("");
			var book = data[0][item].join("");
            // Add schedule content row by row
			sheet.appendRow([lessonNo, hrsDone[item], "", admin, book]);
			lessonNo += 1;
		}
	}
}



function getData(level, sourceHyperlinks) {
	var column2 = [];
	var column3 = [];
	var result = []
	var startColumn = 1;
	// Where the data is comming from. 
	var SPREADSHEET_ID = sourceHyperlinks;
	var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
	var sheet = ss.getSheetByName(level)
	var lastColumn = sheet.getLastColumn() + 1
	var lastRow = sheet.getLastRow() + 1
	var getAdminTask = sheet.getRange(1, 1, 1, lastColumn).getValues();
	var endColumn = getAdminTask[0].indexOf("Admin Task")
    var adminLast = getAdminTask[0].lastIndexOf("Admin Task")
    Logger.log(endColumn);
	//Logger.log(endColumn);
	var adminStart = endColumn
	var startRow = 2;
	//Books array. 
    
    
    
    
    
	while (startRow < lastRow) { 
		column2.push(getRows(sheet, startRow, startColumn, endColumn))
		column3.push(getRows(sheet, startRow, adminStart-1, adminLast-1))
		startRow += 1
	}
	result.push(column2);
	result.push(column3);
	return result
}


function getRows(sheet, startRow, startColumn, endColumn)  {
    var htmlObjects = []
    // Get all the data in one row. 
    var cell = sheet.getRange(startRow, startColumn, 1,  endColumn).getValues();
//    Logger.log(cell)
    var colors = sheet.getRange(startRow, startColumn, 1,  endColumn).getFontColors();
    Logger.log(colors);
    var fontWeight = sheet.getRange(startRow, startColumn, 1,  endColumn).getFontWeights();
    var lineBreak = "";
    var textResult = arrayMaker(cell)
//    Logger.log(textResult);
    var colorResult = arrayMaker(colors)
    var font = arrayMaker(fontWeight)
    for (var i = 1; i < textResult.length; i++) { //textResult.length
        //Logger.log(textResult)
        if (textResult[i][1] === "") {  

          var link = "";
        } else link = "href='" + textResult[i][1] + "' target='_blank'";
          var style = "";
          
          
        if (textResult[i][1] === "") {    

          var textDecoration = "";
        } else textDecoration =" text-decoration:none;";
 
        if (textResult[i][0] === "") {  

          var style= ""
        } else style = " style='color:"+colorResult[i][0]+"; "
        
        if (font[i][0] === "" || font[i][0] === "normal" ) { 

          var fontStyle= "";
        } else var fontStyle = " font-weight:"+ font[i][0];
        
        if (textResult[i][0] === "") {

            lineBreak = "";
        } else lineBreak = "<br>";
        if(textResult[i][0] === ""){
         htmlObjects.push("")
        } else {
        htmlObjects.push("<a"+style+textDecoration+fontStyle+"' "+link+" >"+textResult[i][0]+"</a>"+lineBreak);
        
        }
        
        
    }
    Logger.log(htmlObjects);






    return htmlObjects;
}


// ----------------- Make new schedulefinished -----------------------------------







// ----------------- Standards checker -----------------------------------


function getAllText(itemsToCheck, tabList) {
    Logger.log(itemsToCheck);
    //var itemsToCheck = ['ELT Print certificates'];
	var SHEET_NAME = tabList;
    var lessonNo;
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var output = ss.getSheetByName("Standards Checker");
    var URL;
    if (itemsToCheck.indexOf("URL") >=0){
    URL = true;
    } else {
    URL = false;
    }
	for (var sheet in SHEET_NAME) {
		var sheetToGet = ss.getSheetByName(SHEET_NAME[sheet]);
		var getLastColumn = sheetToGet.getLastColumn();
		var getLastRow = sheetToGet.getLastRow();
		for (var k = 1; k < getLastColumn; k++) {
			for (var i = 1; i < getLastRow + 1; i++) {
				var data = sheetToGet.getRange(i, k).getValues();
				var item = data.toString();
				var conditionResult;
                for (var word in itemsToCheck){
                if (item.indexOf(itemsToCheck[word]) >= 0) {
					conditionResult = true;
                    break;
				} else conditionResult = false;
                }                               
                if (URL === false){
                if (isNaN(item) === true && ValidURL(item) === false && conditionResult === true) {
					lessonNo = sheetToGet.getRange(i, 1).getValues().toString();
					output.appendRow([SHEET_NAME[sheet], item, lessonNo]);
				}
                } else if (URL === true || conditionResult === true){
                if (ValidURL(item) === true) {
					lessonNo = sheetToGet.getRange(i, 1).getValues().toString();
					output.appendRow([SHEET_NAME[sheet], item, lessonNo]);
				}
                }
			}
		}
	}
    tabList.length = 0;
    SHEET_NAME.length = 0; 
    itemsToCheck.length = 0;
    }


// ----------------- Standards checker -----------------------------------








// ----------------- Get Admin tasks -----------------------------------

function getAdminTasks(itemsToCheck, tabList,suffix) {
    Logger.log(itemsToCheck);
    var homework;
    if (itemsToCheck.indexOf("Homework") >=0){
    homework = true;
    } 
    //var itemsToCheck = ['ELT Print certificates'];
	var SHEET_NAME = tabList;
    var lessonNo;
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var output = ss.getSheetByName("Admin Tasks");
	for (var sheet in SHEET_NAME) {
		var sheetToGet = ss.getSheetByName(SHEET_NAME[sheet]);
		var getLastColumn = sheetToGet.getLastColumn();
		var getLastRow = sheetToGet.getLastRow();
        var homeworkLesson = sheetToGet.getRange(getLastRow-1, 1).getValue();
        Logger.log(homeworkLesson);
		for (var k = 1; k < getLastColumn; k++) {
			for (var i = 1; i < getLastRow + 1; i++) {
				var data = sheetToGet.getRange(i, k).getValues();
				var item = data.toString();
				var conditionResult;
                for (var word in itemsToCheck){
                if (item.indexOf(itemsToCheck[word]) >= 0) {
					conditionResult = true;
                    break;
				} else conditionResult = false;
                }                               
                if (isNaN(item) === true && ValidURL(item) === false && conditionResult === true) {
                    
					lessonNo = sheetToGet.getRange(i, 1).getValues().toString();
					output.appendRow([SHEET_NAME[sheet],suffix,"HCM",lessonNo,item,])
				}
               
                }
			}
           /*  if (homework === true){
                output.appendRow([SHEET_NAME[sheet],suffix,"HCM",homeworkLesson,"Homework",])
                }*/
		}
	
    
    tabList.length = 0;
    SHEET_NAME.length = 0; 
    itemsToCheck.length = 0;
    }

// ----------------- Get Adming tasks Finished -----------------------------------







// ----------------- Format test scores for creating template -----------------------------------


function formatTestItem(item){
  var testArr = []
  var removeRevision = item.replace("Revision + ", "").
  replace(" Revision + ", "").
  replace("Revision ", "").
  replace(" Revision ", "").
  replace(" MMT ", "").
  replace("MMT ", "").
  replace("EOMT ", "").
  replace(" EOMT ", "").
  replace("Teacher ", "")
  var testParts = removeRevision.split("+ ")
  for(var part in testParts){
    testArr.push(testParts[part])
  }
  return testArr
}


function cambridgeTestNameFormat(cambridge){
  if(cambridge.indexOf("lesson") < 0) {
    var test = cambridge.search(/\d/)+1
    var testText = cambridge.slice(0,test)  
    return testText;
  } else {
  return "";
  }
}


function cambridgeFormat(cambridge){
  if(cambridge.indexOf("lesson") < 0) {
    var test = cambridge.search(/\d/)+2
    var testText = cambridge.slice(0,test);
    var tests = cambridge.replace(testText,"").split(" + ");  
  
    return tests;
  } else {
  return "";
  }
  
}

// ----------------- Format test scores for creating template Finished -----------------------------------







// ----------------- Create assessment -----------------------------------

          

function getAssessment(itemsToCheck, tabList, suffix) {
    
    
    var ss = SpreadsheetApp.openById('1kGNdPmleXN7p5zZ7j_vpYbTGb4udluZrAbg2ar8i67Y');
	var sheet  = ss.getSheetByName("TabConvertor");
    var levelCodes = sheet.getRange(1,1).getValue();

    
	var homework = true;
	var testGroups = ["Mid Module Test", "End of Module Test", "Supplementary Assessment", "End-of-Project Assessment", "Cambridge"];
	var SHEET_NAME = tabList;
	var lessonNo;
	var testGroup;
	var camTests;
    
    
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var output = ss.getSheetByName("Assessment");
    // Delete and then add header row;
    var header = [
      ["","Template Suffix","ERP Level Code","Module","Lesson No.","Test Group","Test Type","Raw Score","Weight"]
    ];
    output.getRange(1,1,1,9).setValues(header)
	for (var sheet in SHEET_NAME) {
		var sheetToGet = ss.getSheetByName(SHEET_NAME[sheet]);
		var getLastColumn = sheetToGet.getLastColumn();
		var getLastRow = sheetToGet.getLastRow();
		var homeworkLesson = sheetToGet.getRange(getLastRow - 1, 1).getValue(); 
        
        
        //var ERPLevel = levelCodes[SHEET_NAME[sheet]]
         var ERPLevels = JSON.parse(levelCodes);   
         var ERPLevel = ERPLevels[SHEET_NAME[sheet]]
         if(ERPLevel === undefined) return ["System can't read the tab. \n Did you update the tab convertion table?","https://docs.google.com/spreadsheets/d/1kGNdPmleXN7p5zZ7j_vpYbTGb4udluZrAbg2ar8i67Y/edit#gid=1863587942"];
        
		Logger.log(homeworkLesson);
		for (var k = 1; k < getLastColumn; k++) {
			var projectNum = 1;
			for (var i = 2; i < getLastRow + 1; i++) {
				var data = sheetToGet.getRange(i, k).getValues();
				var item = data.toString();
				//         for (var test in itemsToCheck) {
				if (item.indexOf("Cambridge") >= 0 && item.indexOf("Lesson") < 0 && item.indexOf("http") < 0 && item.indexOf("Letter") < 0) {
					testGroup = cambridgeTestNameFormat(item)
					camTests = cambridgeFormat(item)
					for (var z = 0; z < camTests.length; z++) {
						lessonNo = sheetToGet.getRange(i, 1).getValues().toString();
                        
						output.appendRow([SHEET_NAME[sheet], suffix, ERPLevel[1],ERPLevel[0], lessonNo, testGroup, camTests[z]])
                      //  return ([SHEET_NAME[sheet]] + " This is the ERP Level: "+ ERPLevel[1]); // testSheetName
                        
					}
				}
				if (item.indexOf("MMT") >= 0 || item.indexOf("EOMT") >= 0 || item.indexOf("End-of-Project Assessment") >= 0 && item.indexOf("http") < 0) {
					Logger.log(item)
					debugger;
					var tests = formatTestItem(item)
					if (item.indexOf("MMT") >= 0) {
						testGroup = testGroups[0];
					} else if (item.indexOf("EOMT") >= 0) {
						testGroup = testGroups[1];
					} else if (item.indexOf("End-of-Project Assessment") >= 0) {
						testGroup = testGroups[3];
						tests = ["Project " + projectNum];
						projectNum++;
					}
					for (var x = 0; x < tests.length; x++) {
						lessonNo = sheetToGet.getRange(i, 1).getValues().toString();
						output.appendRow([SHEET_NAME[sheet], suffix, ERPLevel[1],ERPLevel[0], lessonNo, testGroup, tests[x]])
                     //   return ([SHEET_NAME[sheet]] + " This is the ERP Level: "+ ERPLevel[1]); // testSheetName
					}
				}
			}
		}
                    var supplimentry = searchSupp(SHEET_NAME[sheet]);
                    if(supplimentry[2] === "FAILED") return supplimentry;
                    for(var j = 0; j < supplimentry.length; j++){
                      output.appendRow([SHEET_NAME[sheet], suffix, ERPLevel[1],ERPLevel[0], homeworkLesson, testGroups[2], supplimentry[j]]);                        
                   //   return ([SHEET_NAME[sheet]] + " This is the ERP Level: "+ ERPLevel[1]);  // testSheetName                      
                    }
        
                
		
	}
	tabList.length = 0;
	SHEET_NAME.length = 0;
	itemsToCheck.length = 0;
}


//****************** Unit test for matching Sheet Tab names with ERP Level Names *****************************
// Data course for tab/level converstion https://docs.google.com/spreadsheets/d/1kGNdPmleXN7p5zZ7j_vpYbTGb4udluZrAbg2ar8i67Y/edit#gid=1863587942
function testSheetName(){
    var itemsToCheck =["String"];
    var tabList = ['J1B'];
    var suffix = "String"
    var expected = "System can't read the tab. Did you update the tab in the tab convertion table?";
    assert(getAssessment(itemsToCheck, tabList, suffix),expected,"It checks sheet names are being processed");
}

function testSupplimetry(){
    var itemsToCheck =["String"];
    var tabList = ['J1B'];
    var suffix = "String"
    var expected = JSON.stringify(["Failed to find this tab name/level \n in the Supplimentry Assessment document","https://docs.google.com/spreadsheets/d/1PIMtKuiA9hzxb-9GT0dE7Wvn-4pZC8nrOx_943dk8YQ/edit#gid=465027747"]);
    assert(JSON.stringify(getAssessment(itemsToCheck, tabList, suffix)),expected,"It checks if warning array is returned");
}







// ----------------- Create assessment finished -----------------------------------

// Supplimentary assessment
function  searchSupp(level)  {
        var result = []
       // var level = "S-A1"
        var SPREADSHEET_ID = "1PIMtKuiA9hzxb-9GT0dE7Wvn-4pZC8nrOx_943dk8YQ"
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName("Supplimentary Assessment")
        var maxRows = sheet.getMaxRows();
        var maxColumns = sheet.getMaxColumns()-1
        var levelRow = sheet.getRange(1,1,maxRows).getValues();
        var levels = [] // [first row, second Row]
            levelRow.map(function(item, index){
              if(item[0] === level){
                levels.push(Number(index)+1);
              }
            })
        var warningText = "Failed to find this tab name/level in the \n Supplimentry Assessment document";    
        var warningURL = "https://docs.google.com/spreadsheets/d/1PIMtKuiA9hzxb-9GT0dE7Wvn-4pZC8nrOx_943dk8YQ/edit#gid=465027747"
        if(levels.length === 0) return [warningText, warningURL,"FAILED"];
        
        var yes = sheet.getRange(levels[0],2,1,maxColumns).getValues();
        var headers = sheet.getRange(1,2,1,maxColumns).getValues();
        yes[0].forEach(function(item, ind){
            if(item === "yes"){
            result.push(headers[0][ind]);
            }
        })
        return result;
}








// ----------------- Call to technical scores template to get assessment scores --------------------

function scoreRunner(sheet){
        Logger.log(sheet);
        var scores;
        var target = SpreadsheetApp.getActiveSpreadsheet();
	    var targetSheet = target.getSheetByName("Assessment");
        var maxRows = targetSheet.getMaxRows()+1;               
        for (var i = 2; i<maxRows; i++){
        var rowOne = targetSheet.getRange(i, 1,1,7).getValues();
        var testGroup = rowOne[0][5].trim()
        var testType = rowOne[0][6].trim()
        var level = rowOne[0][0]
        
        Logger.log(testGroup);
        if (testGroup.indexOf("Cambridge") >=0){
        var scores = getCambridgeScores(level,testGroup,testType)
        } else {
        var scores = searchScores(level,testGroup,testType, sheet)
        }
        
        targetSheet.getRange(i, 8).setValue(scores[1]);
        targetSheet.getRange(i, 9).setValue(scores[0]);
        }
        
}


function  searchScores(searchTerm,testGroup,testType, sheet)  {
        var searchTerm = searchTerm
        var SPREADSHEET_ID = "1PIMtKuiA9hzxb-9GT0dE7Wvn-4pZC8nrOx_943dk8YQ"
        var locatedCells = [];
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(sheet)
        Logger.log(sheet)
        // Get rows for Levels 
        var maxRows = sheet.getMaxRows()
        var maxColumns = sheet.getMaxColumns()-1
        var levelRow = sheet.getRange(1,2,maxRows).getValues();
        var levels = [] // [first row, second Row]
            levelRow.map(function(item, index){
              if(item[0] === searchTerm){
                levels.push(Number(index)+1);
              }
            })
        // Get column for Test item. 
         var column = 0;
         var testGroupCol = sheet.getRange(1,1,1,maxColumns).getValues();
         var testTypeCol = sheet.getRange(2,1,1,maxColumns).getValues();
                  for (var i = 0; i < testTypeCol[0].length; i++){              
                if((testGroupCol[0][i] === testGroup || testGroup.indexOf("Cambridge") >=0)  && testTypeCol[0][i] === testType ){
                column = i+1;
                }        
            }        
        var weight = sheet.getRange(levels[0], column).getValue()
        var score = sheet.getRange(levels[1], column).getValue()
        return [weight, score]
        
        
}


function getCambridgeScores(level,testGroup,testType){     
     var sheet = "Cambridge Raw Scores"
     var result = searchScores(level,testGroup,testType, sheet)
        return result
}

// ----------------- Call to technical scores template to get assessment scores finished -----------------------------------        


//___________________ Copy Admin Tasks to ERP Template ---------------------------------------------------------------------

function copyAdmin(){
 var source = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = source.getSheetByName("Admin Tasks");
 var destination = SpreadsheetApp.openById('16aVz_mBKfmPA4hCkaJ4HEkJy5fnprojkLi-a6NGO0KU');
 var destSheet = destination.getSheetByName("AdminTasks");
 var data = sheet.getDataRange().getValues();
 Logger.log(data);
 for(var row in data){
  destSheet.appendRow(data[row])
 }  
}





////___________________ Copy Admin Tasks to ERP Template ---------------------------------------------------------------------



    

    
 
 
 
 

//function search(SPREADSHEET_ID, SHEET_NAME, searchTerm) {
//        var locatedCells = [];
//        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
//        var output = ss.getSheetByName("Standards Checker")
//        var searchLocation = ss.getSheetByName(SHEET_NAME).getDataRange().getValues();        
//        //Loops to find the search term. 
//            for (var j = 0, jLen = searchLocation.length; j < jLen; j++) {
//                for (var k = 0, kLen = searchLocation.length; k < kLen; k++) {
//                    var find = searchTerm;
//                    if (find == searchLocation[j][k]) {
//                        
//                         output.appendRow([SHEET_NAME,searchTerm,j + 1,k + 1])                   }
//                }
//                
//               
//            }
//         //   Logger.log(locatedCells);
//       //     return(locatedCells)
//        }
////=QUERY(A:A,"SELECT COUNT(A) WHERE A ='"&E3&"'LABEL COUNT(A)''",0)




