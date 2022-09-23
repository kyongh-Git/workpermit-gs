// Work permit google apps script
function doGet(e){
  return handleResponse(e);
}

//  Enter sheet name where data is to be written below
        var SHEET_NAME = "Sheet1";

//var SCRIPT_PROP = PropertiesService.getScriptProperties(); // new property service

function handleResponse(e) {
  var lock = LockService.getPublicLock();
  lock.waitLock(30000);
  
  var std_cwid = "";
  
  try {
    // next set where we write the data - you could write to multiple/alternate destinations
    var doc = SpreadsheetApp.openById("1g0PGPQVNUsh9aLNAUMPuFKi1NXv1GE9uriAFalQLB98");//SCRIPT_PROP.getProperty("key")
    var sheet = doc.getSheetByName(SHEET_NAME);
    
    // we'll assume header is in row 1 but you can override with header_row in GET/POST data
    var headRow = e.parameter.header_row || 1;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var nextRow = sheet.getLastRow()+1; // get next row
    var row = []; 
    // loop through the header columns
   for (var i = 0; i < headers.length; i++) { // start at 1 to avoid Timestamp column
            if (headers[i].length > 0) {
                row.push(e.parameter[headers[i]]); // add data to row
            }
        }
            var cwid = String(row[0]);
            std_cwid = cwid;
            var address = String(row[1]);
            var phone = String(row[2]);
            var entry = String(row[3]);
            var position = String(row[4]);
            var coc = String(row[5]);
            var cor = String(row[6]);
            var stat = String(row[7]);
            var level = String(row[8]);
            var hrsEnrol = String(row[9]);
            var sem = String(row[10]);
            var sig = String(row[11]);
            var fn = String(row[12]);
            var ln = String(row[13]);
            var stdEmail = String(row[14]);
            var emptwent = String(row[15]);  //
            var empfull = String(row[16]);   //
            var effdafrom = String(row[17]);
            var effdato = String(row[18]);
            var stt = String(row[19]);
            var jcate = String(row[20]);
            var time = String(row[21]);
            var fica = String(row[22]);
            var nonali = String(row[23]);
            var cocd = String(row[24]);
            var expries = String(row[25]);
            var comm = String(row[26]);
            var appr = String(row[27]);
            var street = String(row[28]);
            var add2 = String(row[29]);
            var city = String(row[30]);
            var state = String(row[31]);
            var zip = String(row[32]);
            var country = String(row[33]);
            var wpstdate = String(row[34]);
            
            //var today = Utilities.formatDate(new Date(), "GMT+1", "MM/dd/yyyy");
            
                    var name = cwid + " Work Permit Form";
                    traveltemp = DriveApp.getFileById('1uQEHm0lGyMFuYm4yjRhLeQl8kxaY7mLEY7LLdFCdKEg').makeCopy(name);
                    var id = traveltemp.getId();
                    var worddoc = DocumentApp.openById(id);
                    var body = worddoc.getBody();

                    body.replaceText("<lname>", ln);
                    body.replaceText("<fname>", fn);
                    body.replaceText("<cwid>", cwid);
                    body.replaceText("<address>", street + ", " + city + ", " + state + ", " + zip + ", " + country);
                    body.replaceText("<phoneNum>", phone);
                    body.replaceText("<entDate>", entry);
                    body.replaceText("<position>", position);
                    body.replaceText("<sig>", sig);
                    body.replaceText("<coc>", coc);
                    body.replaceText("<cor>", cor);
                    body.replaceText("<today>", wpstdate);
                    body.replaceText("<stat>", stat);
                    body.replaceText("<level>", level);
                    body.replaceText("<hrEn>", hrsEnrol);
                    body.replaceText("<sem>", sem);
                    body.replaceText("<from>", effdafrom);
                    body.replaceText("<to>", effdato);
                    body.replaceText("<stt>", stt);                    
                    body.replaceText("<to>", effdato);
                    body.replaceText("<time>", time);
                    body.replaceText("<fica>", fica);
                    body.replaceText("<nonali>", nonali);
                    body.replaceText("<cocd>", cocd);
                    body.replaceText("<expdate>", expries);
                    
                    body.replaceText("<apprv>", appr);
                   
                    if(jcate == undefined)
                    {
                        body.replaceText("<jcate>", ""); 
                    }
                    else
                    {
                        body.replaceText("<jcate>", jcate); 
                    }
                    if(jcate == undefined)
                    {
                        body.replaceText("<comm>", "")
                    }
                    else
                    {
                        body.replaceText("<comm>", comm)
                    }
                    
                    var checkbox_image = DriveApp.getFileById('1OIrVi3FmK8o9YPjV8q5nA6Lh44C5MAXn').getBlob();
                    var box_image = DriveApp.getFileById('1je-HuiZjW1ZF8olropFKAgHNlRiJLOlx').getBlob();
                    
                    if(emptwent == "Yes")
                    {
                        var next = body.findText("<check1>");
                        var r = next.getElement();
                        r.asText().setText("");
                        var img = r.getParent().asParagraph().insertInlineImage(0, checkbox_image);
                    }
                    else if(emptwent == "No")
                    {
                        var next = body.findText("<check1>");
                        var r = next.getElement();
                        r.asText().setText("");
                        var img = r.getParent().asParagraph().insertInlineImage(0, box_image);                    
                    }
                    else
                    {
                        var next = body.findText("<check1>");
                        var r = next.getElement();
                        r.asText().setText("");
                        var img = r.getParent().asParagraph().insertInlineImage(0, box_image);  
                    }
                    
                    if(empfull == "Yes")
                    {
                        var next = body.findText("<check2>");
                        var r = next.getElement();
                        r.asText().setText("");
                        var img = r.getParent().asParagraph().insertInlineImage(0, checkbox_image);
                    }
                    else if(empfull == "No")
                    {
                        var next = body.findText("<check2>");
                        var r = next.getElement();
                        r.asText().setText("");
                        var img = r.getParent().asParagraph().insertInlineImage(0, box_image);                    
                    }
                    else
                    {
                        var next = body.findText("<check2>");
                        var r = next.getElement();
                        r.asText().setText("");
                        var img = r.getParent().asParagraph().insertInlineImage(0, box_image);  
                    }
                  
                    worddoc.saveAndClose();
                    row.push(traveltemp.getUrl());
               
              
               
    // more efficient to set values as [][] array than individually
    sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);
    //Prepare and send emails
                    var approvalEmail = "test@gmail.com";
                    var subjectLine = "Work permit form of " + cwid + " " + fn + " " + ln;
                    MailApp.sendEmail(approvalEmail, subjectLine, "The student's work permit form is attached",{
                        cc: stdEmail,
                        name: "Work Permit Form",
                        attachments: [traveltemp.getAs(MimeType.PDF)]  
                    });
    // return json success results
    return ContentService
          .createTextOutput(JSON.stringify({"result":"success", "row": nextRow})) 
          .setMimeType(ContentService.MimeType.JSON);
  } catch(e){
    // if error return this
     var email = "yjckimdyd@gmail.com";
                    var subjectLine = "Work permit Request failed for " + std_cwid ;
                    var body = "<HTML><BODY>" + e.message + "</BODY></HTML>";
                    MailApp.sendEmail({
                        to: email,
                        subject: subjectLine,
                        htmlBody: body,

                        name: "Work Permit",
                        attachments: []
                    });
    return ContentService
          .createTextOutput(JSON.stringify({"result":"error", "error": e}))
          .setMimeType(ContentService.MimeType.JSON);
  } finally { //release lock
  }
}

function setup() {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    SCRIPT_PROP.setProperty("key", doc.getId());
}