/*
  ++++++++++++++++++++++++++++++++++++++++++++++++
  + Google docs, automated report script         +
  + Author: Carlos Camacho <ccamacho@redhat.com> +
  ++++++++++++++++++++++++++++++++++++++++++++++++

  README:
  Fill the variables below the multiline comment:
*/

/* Trello variables */
var TRELLO_KEY = '<trello_key_goes_here>';
var TRELLO_TOKEN = '<trello_token_goes_here>';
var TRELLO_LIST_ID = ["<List_1_ID>", "<List_2_ID>","<List_3_ID>"]; // Add as many lists as you need
var TRELLO_TITLES = ["<List_1_Title>", "<List_2_Title>","<List_3_Title>"]; // Add as meny titles as lists
var TRELLO_USER_FILTER = ["<Person_name_to_filter_cards>", "<Person_name_to_filter_cards>"]; // Only display these people cards

/* Stackalytics variables */
var STACKALYTICS_USER= "<stackalytics_user>";
/* Bugzilla variables */
var BZ_HOST = "https://bugzilla.redhat.com";
var BZ_STATUS = "bug_status=NEW&bug_status=ASSIGNED&bug_status=POST&bug_status=MODIFIED&bug_status=ON_DEV&bug_status=ON_QA&bug_status=VERIFIED&bug_status=RELEASE_PENDING";
var BZ_EMAIL = "<bugzilla_email>"; // Like: user%40domain.com";
/* Launchpad variables */
var LAUNCHPAD_USER = "<launchpad user>";
/* Storyboard variables */
var STORYBOARD_USER = "<Storyboard_user_id>";


/*
  DO NOT EDIT BELOW THIS COMMENT OR BAD THINGS WILL HAPPEN
*/

NUMBER_OF_LINES = 51

Date.prototype.getUnixTime = function() { return this.getTime()/1000|0 };
if(!Date.now) Date.now = function() { return new Date(); }
Date.time = function() { return Date.now().getUnixTime(); }


Date.prototype.addDays = function(days) {
  var dat = new Date(this.valueOf());
  dat.setDate(dat.getDate() + days);
  return dat;
}

function getNextValidDateFrom(addDay) {
  var nextValid = new Date().addDays(addDay);
  return nextValid.toDateString();
}
  
function insertDateAndTable(date, line) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var cursor = doc.getCursor();
  if (cursor) {
    var element = cursor.getElement();
    var idx = element.getParent().getChildIndex(element) + 1;
    var heading = body.insertParagraph(idx, 'Agenda for ' + date + "\n");
    heading.setHeading(DocumentApp.ParagraphHeading.HEADING1)
        .setBold(true);    
    var table = body.insertTable(idx+1);
    var tr = table.appendTableRow();
    var number_str = "0\n"
    for(var j=1; j<line; j++){
      number_str += j + "\n";
    }
    var td = tr.appendTableCell(number_str);
    var td = tr.appendTableCell('');
    table.setColumnWidth(0,30);
    table.setBorderWidth(0.5);
    table.setBorderColor('#d3d3d3');
  } else {
    DocumentApp.getUi().alert('Cannot find a cursor.');
  }
}

function createTodayAgenda() {
  var date = getNextValidDateFrom(0);
  insertDateAndTable(date, NUMBER_OF_LINES)
}
function createTomorrowAgenda() {
  var date = getNextValidDateFrom(1);
  insertDateAndTable(date, NUMBER_OF_LINES)
}

//RELATIVE TO THE AGENDA ITEMS CREATION.

function insertAgenda(agenda, name) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var cursor = doc.getCursor();
  if (cursor) {
    var element = cursor.getElement();
    var idx = element.getParent().getChildIndex(element);
    idx = idx + 1;
    var heading = body.insertParagraph(idx, "  " + name + " - Action items:")
          .setFontSize(16);
    
    
    
    //heading.setBold(true);
    for (var person in agenda) {
      if (TRELLO_USER_FILTER.indexOf(person.replace(/["']/g, "")) > -1){

        idx = idx + 1; 
        var actionPerson = "    ["+person.replace(/["']/g, "")+"]:";
        var heading = body.insertParagraph(idx, actionPerson)
            .setFontSize(12)
            .setBold(false);
        for (var task in agenda[person]) {
        idx = idx + 1; 
          var actionItem = agenda[person][task]['title'].replace(/["]/g, "");
          var labels = JSON.parse(agenda[person][task]['labels']);
          var desc = JSON.parse(agenda[person][task]['desc']);
          var labels_array = [];      
          for (var index in labels) {
            labels_array.push(labels[index]['name']);
          }
          var line = "      * "+actionItem+" "+" ["+labels_array.join(',')+"] "
          var heading = body.insertParagraph(idx, line)
            .setFontSize(10)
            .setBold(true);
          var parr = heading.appendText("Link ");
          parr.setLinkUrl(agenda[person][task]['url'].replace(/["']/g, ""));
          idx = idx + 1; 
          var description_body = body.insertParagraph(idx, desc)
            .setFontSize(10)
            .setIndentEnd(30)
            .setIndentFirstLine(30)	
            .setIndentStart(30)
            .setBold(false);
        }
      }
    }
    var empty = body.insertParagraph(idx + 1, "  ");    
  } else {
    DocumentApp.getUi().alert('Cannot find a cursor.');
  }  
}

function insertHeader(date) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var cursor = doc.getCursor();
  if (cursor) {
    var element = cursor.getElement();
    var idx = element.getParent().getChildIndex(element);
    var heading = body.insertParagraph(idx, 'Agenda for ' + date);
    heading.setHeading(DocumentApp.ParagraphHeading.HEADING1)
        .setBold(true);
    idx = idx + 1;
  } else {
    DocumentApp.getUi().alert('Cannot find a cursor.');
  }  
}

function formatDate(date) {
  return date.getFullYear() + "/" + (+date.getMonth()+1) + "/" + date.getDate();
}

function insertReportHeader(date1, date2) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var cursor = doc.getCursor();
  if (cursor) {
    var element = cursor.getElement();
    var idx = element.getParent().getChildIndex(element);
    var heading = body.insertParagraph(idx, 'Activity report for: ' + formatDate(date1) + ' - ' + formatDate(date2));

    heading.setHeading(DocumentApp.ParagraphHeading.HEADING1)
        .setBold(true);    
    idx = idx + 1;
  } else {
    DocumentApp.getUi().alert('Cannot find a cursor.');
  }  
}

function insertNotes() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var cursor = doc.getCursor();
  if (cursor) {
    var element = cursor.getElement();
    var idx = element.getParent().getChildIndex(element);
    idx = idx + 1; 
    var heading = body.insertParagraph(idx, "  Additional notes:")
          .setFontSize(16)
          .setBold(true);
    idx = idx + 1;
    var table = body.insertTable(idx)
          .setFontSize(10)
          .setBold(false);
    var tr = table.appendTableRow();
    var number_str = "0\n"
    var additional_lines = 10+1;
    for(var j=1; j<additional_lines; j++){
      number_str += j + "\n";
    }
    var td = tr.appendTableCell(number_str);
    var td = tr.appendTableCell('');
    table.setColumnWidth(0,30);
    table.setBorderWidth(0.5);
    table.setBorderColor('#d3d3d3');
  } else {
    DocumentApp.getUi().alert('Cannot find a cursor.');
  }  
}

function createItemsForAgenda() {
  var date = getNextValidDateFrom(0);
  insertHeader(date);
  insertNotes();
  for (var i = TRELLO_LIST_ID.length - 1; i >= 0; --i) {;
    insertAgenda(getListDetails(TRELLO_LIST_ID[i]), TRELLO_TITLES[i]);
  }
}

function insertGerritReport(date1, date2) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var cursor = doc.getCursor();
  if (cursor) {
    var element = cursor.getElement();
    var idx = element.getParent().getChildIndex(element);
    
    var url  = "http://stackalytics.com/api/1.0/contribution?release=all&user_id=" + STACKALYTICS_USER + "&start_date=" + date1.getUnixTime() + "&end_date="+date2.getUnixTime() + "&page_size=1000000"; 
    var options = {
      "method" : "GET"
    };
    try {
      var response = UrlFetchApp.fetch(url, options);
    }catch(err) {
      var ui = DocumentApp.getUi();
      var result = ui.alert('Error when fetching the Stackalytics URL:', JSON.stringify(err.message), ui.ButtonSet.OK);
      return {};
    }
    var data = JSON.parse(response.getContentText());
   
    for (var index in data['contribution']) {
      if(index == 'commit_count' || index == 'translations' || index == 'loc' || index == 'completed_blueprint_count' || index == 'drafted_blueprint_count' || index == 'resolved_bug_count'){
        var description_body = body.insertParagraph(idx, index + ": " + data['contribution'][index])
            .setFontSize(10)
            .setIndentEnd(30)
            .setIndentFirstLine(30)	
            .setIndentStart(30)
            .setBold(false);
        idx = idx + 1;    
      }
      if(index == 'marks'){
        var description_body = body.insertParagraph(idx, 'Reviews:')
            .setFontSize(10)
            .setIndentEnd(30)
            .setIndentFirstLine(30)	
            .setIndentStart(30)
            .setBold(false);
          idx = idx + 1;            
        
        for (var index in data['contribution']['marks']) {
          var review_title = '';
          switch(index){
            case '0':
              review_title = 'Code-Review 0';
              break;
            case '1':
              review_title = 'Code-Review +1';
              break;
            case '2':
              review_title = 'Code-Review +2';
              break;
            case '-1':
              review_title = 'Code-Review -1';
              break;              
            case '-2':
              review_title = 'Code-Review -2';
              break;
            case 'A':
              review_title = 'Approved';
              break;
            case 'WIP':
              review_title = 'Work in progress';
              break;              
            case 's':
              review_title = 'Self review';
              break;
            case 'x':
              review_title = "Don't know what this is";
              break;
          }
          
          var description_body = body.insertParagraph(idx, "  " + review_title + ": " + data['contribution']['marks'][index])
            .setFontSize(10)
            .setIndentEnd(30)
            .setIndentFirstLine(30)	
            .setIndentStart(30)
            .setBold(false);
          idx = idx + 1;              
        }
      }      
    }        
  } else {
    DocumentApp.getUi().alert('Cannot find a cursor.');
  }  
}

function insertBugzillaReport(date1, date2) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var cursor = doc.getCursor();
  if (cursor) {
    var element = cursor.getElement();
    var idx = element.getParent().getChildIndex(element);

    var isodate1 = date1.getFullYear() + "-" + (date1.getMonth() + 1) + "-" + date1.getDate();
    var isodate2 = date2.getFullYear() + "-" + (date2.getMonth() + 1) + "-" + date2.getDate();
    
    var queryformat = "&emailassigned_to1=1&emailtype1=exact&list_id=8803296&query_format=advanced&title=Bug%20List&ctype=atom";
    
    var url  = BZ_HOST + "/buglist.cgi?" + BZ_STATUS + "&chfieldfrom=" + isodate1 + "&chfieldto=" + isodate2 + "&email1=" + BZ_EMAIL + queryformat; 
    var options = {
      "method" : "GET"
    };
    try {
      var response = UrlFetchApp.fetch(url, options);
    }catch(err) {
      var ui = DocumentApp.getUi();
      var result = ui.alert('Error when fetching the Bugzilla URL:', JSON.stringify(err.message), ui.ButtonSet.OK);
      return {};
    }

    var atom = XmlService.getNamespace('http://www.w3.org/2005/Atom');
    var rawbz = XmlService.parse(response);

    var entries = rawbz.getRootElement().getChildren('entry', atom);
    
    for (var i = 0; i < entries.length; i++) {
        var title = entries[i].getChild('title', atom).getText();
        var id = entries[i].getChild('id', atom).getText();

        var heading = body.insertParagraph(idx, '    ')
          .setFontSize(10);
        var parr = heading.appendText(title);
        parr.setLinkUrl(id);
        idx = idx + 1;
    }
  } else {
    DocumentApp.getUi().alert('Cannot find a cursor.');
  }  
}

function myAtomDateParser(datestr) {
var yy   = datestr.substring(0,4);
var mo   = datestr.substring(5,7);
var dd   = datestr.substring(8,10);
var hh   = datestr.substring(11,13);
var mi   = datestr.substring(14,16);
var ss   = datestr.substring(17,19);
var tzs  = datestr.substring(19,20);
var tzhh = datestr.substring(20,22);
var tzmi = datestr.substring(23,25);
var myutc = Date.UTC(yy-0,mo-1,dd-0,hh-0,mi-0,ss-0);
var tzos = (tzs+(tzhh * 60 + tzmi * 1)) * 60000;
return new Date(myutc-tzos);
}

function insertLaunchpadReport(date1, date2) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var cursor = doc.getCursor();
  if (cursor) {
    var element = cursor.getElement();
    var idx = element.getParent().getChildIndex(element);
    var url  = "http://feeds.launchpad.net/~" + LAUNCHPAD_USER + "/latest-bugs.atom"; 
    var options = {
      "method" : "GET"
    };
    try {
      var response = UrlFetchApp.fetch(url, options);
    }catch(err) {
      var ui = DocumentApp.getUi();
      var result = ui.alert('Error when fetching the Launchpad URL:', JSON.stringify(err.message), ui.ButtonSet.OK);
      return {};
    }

    var atom = XmlService.getNamespace('http://www.w3.org/2005/Atom');
    var rawbz = XmlService.parse(response);
    var entries = rawbz.getRootElement().getChildren('entry', atom);
    
    for (var i = 0; i < entries.length; i++) {
        var title = entries[i].getChild('title', atom).getText();
        var id = entries[i].getChild('link', atom).getAttribute('href');
        var published = entries[i].getChild('published', atom).getText();
        var pdate = myAtomDateParser(published);
        var link = id.getValue();
      if (pdate <= date2 && pdate >= date1){      
        var heading = body.insertParagraph(idx, '    ')
          .setFontSize(10);
        var parr = heading.appendText(title);
        parr.setLinkUrl(link);
        idx = idx + 1;
      }
    }
  } else {
    DocumentApp.getUi().alert('Cannot find a cursor.');
  }  
}

function insertStoryboardReport(date1, date2) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var cursor = doc.getCursor();
  if (cursor) {
    var element = cursor.getElement();
    var idx = element.getParent().getChildIndex(element);

    var url  = "https://storyboard.openstack.org/api/v1/stories?assignee_id=" + STORYBOARD_USER; 
    var options = {
      "method" : "GET"
    };
    try {
      var response = UrlFetchApp.fetch(url, options);
    }catch(err) {
      var ui = DocumentApp.getUi();
      var result = ui.alert('Error when fetching the Storyboard URL:', JSON.stringify(err.message), ui.ButtonSet.OK);
      return {};
    }
    var entries = JSON.parse(response.getContentText());
    for (var i = 0; i < entries.length; i++) {
      var title = entries[i]['title'];
      var id = entries[i]['id'];
      var published = entries[i]['created_at'];
      var pdate = myAtomDateParser(published);
      var link = "https://storyboard.openstack.org/#!/story/" + id;
      if (pdate <= date2 && pdate >= date1){      
        var heading = body.insertParagraph(idx, '    ')
          .setFontSize(10);
        var parr = heading.appendText(title);
        parr.setLinkUrl(link);
        idx = idx + 1;
      }
    }
  } else {
    DocumentApp.getUi().alert('Cannot find a cursor.');
  }  
}

function insertGerritTitle() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var cursor = doc.getCursor();
  if (cursor) {
    var element = cursor.getElement();
    var idx = element.getParent().getChildIndex(element);
    var heading = body.insertParagraph(idx, '  Gerrit report:')
        .setFontSize(16)
        .setBold(true);    
    idx = idx + 1;
  } else {
    DocumentApp.getUi().alert('Cannot find a cursor.');
  }  
}

function insertBugzillaTitle() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var cursor = doc.getCursor();
  if (cursor) {
    var element = cursor.getElement();
    var idx = element.getParent().getChildIndex(element);
    var heading = body.insertParagraph(idx, '  Bugzilla report:')
        .setFontSize(16)
        .setBold(true);    
    idx = idx + 1;
  } else {
    DocumentApp.getUi().alert('Cannot find a cursor.');
  }  
}

function insertLaunchpadTitle() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var cursor = doc.getCursor();
  if (cursor) {
    var element = cursor.getElement();
    var idx = element.getParent().getChildIndex(element);
    var heading = body.insertParagraph(idx, '  Launchpad report:')
        .setFontSize(16)
        .setBold(true);    
    idx = idx + 1;
  } else {
    DocumentApp.getUi().alert('Cannot find a cursor.');
  }  
}

function insertStoryboardTitle() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var cursor = doc.getCursor();
  if (cursor) {
    var element = cursor.getElement();
    var idx = element.getParent().getChildIndex(element);
    var heading = body.insertParagraph(idx, '  Storyboard report:')
        .setFontSize(16)
        .setBold(true);    
    idx = idx + 1;
  } else {
    DocumentApp.getUi().alert('Cannot find a cursor.');
  }  
}

function createActivityReport() {
  var date = getNextValidDateFrom(0);
  var d1, d2;
  var ui = DocumentApp.getUi();

  var htmlDlg = HtmlService.createHtmlOutputFromFile('HTML_myHtml')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(200)
      .setHeight(150);
  
  var ui = DocumentApp.getUi()
  .showModalDialog(htmlDlg, 'Select the quarter to generate the activity report:');    
}

function createActReport(bigdate) {
  var res = bigdate.split(",");
  createFullReport(res[0],res[1]);
}
  
function createFullReport(date1, date2) {  
  d1=new Date(date1);
  d2=new Date(date2);

  insertNotes();
  insertGerritReport(d1, d2);
  insertGerritTitle();
  insertBugzillaReport(d1, d2);
  insertBugzillaTitle();
  insertLaunchpadReport(d1, d2);
  insertLaunchpadTitle();
  insertStoryboardReport(d1, d2);
  insertStoryboardTitle();
  insertReportHeader(d1, d2);
  
}

function getListDetails(listid){
  var agenda = {};
  var url  = "https://api.trello.com/1/list/" + listid + "/cards" + "?key="+ TRELLO_KEY +"&token=" + TRELLO_TOKEN; 
  var options = {
    "method" : "GET"
  };
  try {
    var response = UrlFetchApp.fetch(url, options);
  }catch(err) {
    var ui = DocumentApp.getUi();
    var result = ui.alert('Error when fetching the Trello URL:', JSON.stringify(err.message), ui.ButtonSet.OK);
    return {};
  }
  var data = JSON.parse(response.getContentText());
  var aux = "";
  for (var index in data) {
    var cardurl  = "https://api.trello.com/1/cards/" + JSON.stringify(data[index]['id']).replace(/["']/g, "") + "?key="+ TRELLO_KEY +"&token=" + TRELLO_TOKEN; 
    try {
      var cardraw  = UrlFetchApp.fetch(cardurl, options);
    }catch(err) {
      var ui = DocumentApp.getUi();
      var result = ui.alert('Error when fetching the Trello URL:', JSON.stringify(err.message), ui.ButtonSet.OK);
      return {};
    }  
    var carddata = JSON.parse(cardraw);
    aux = aux + "Card Name: " + JSON.stringify(carddata['name']) + "\n";
    aux = aux + "Card URL: " + JSON.stringify(carddata['url']) + "\n";
    aux = aux + "Card tags: " + JSON.stringify(carddata['labels']) + "\n";
    for (var index in carddata['idMembers']) {
      if(carddata['idMembers'].length > 0){
        carddata['idMembers'][index]
        var userurl  = "https://api.trello.com/1/members/" + JSON.stringify(carddata['idMembers'][index]).replace(/["']/g, "") + "?key="+ TRELLO_KEY +"&token=" + TRELLO_TOKEN; 
        try {
          var userraw  =  UrlFetchApp.fetch(userurl, options);
        }catch(err) {
          var ui = DocumentApp.getUi();
          var result = ui.alert('Error when fetching the Trello URL:', JSON.stringify(err.message), ui.ButtonSet.OK);
          return {};
        }          
        var userdata = JSON.parse(userraw);          
        aux = aux + "Member user: " + JSON.stringify(userdata['fullName']) + "\n";
        if (typeof agenda[JSON.stringify(userdata['fullName'])] !== 'undefined') {
          agenda[JSON.stringify(userdata['fullName'])].push({"title":JSON.stringify(carddata['name']) , "url":JSON.stringify(carddata['shortUrl']), "labels":JSON.stringify(carddata['labels']), "desc":JSON.stringify(carddata['desc'])});
        }else{
          agenda[JSON.stringify(userdata['fullName'])] = new Array({"title":JSON.stringify(carddata['name']) , "url":JSON.stringify(carddata['shortUrl']), "labels":JSON.stringify(carddata['labels']), "desc":JSON.stringify(carddata['desc'])});
        }
      }
    }
    aux = aux + "--------" + "\n";
  }
  return agenda;
}

function onOpen() {
  var ui = DocumentApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Scrum')
    .addItem('Create today\'s agenda', 'createTodayAgenda')
    .addItem('Create tomorrow\'s agenda', 'createTomorrowAgenda')
    .addItem('Create today\'s agenda with Trello items', 'createItemsForAgenda')
    .addItem('Create activity report', 'createActivityReport')

    .addSeparator()
    .addToUi();
}
