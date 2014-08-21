// global configuration
var Conf = {
  sheet_name: "資料",               // pattern of sheet name to check
  tmpl_name: {
    finished: "[email樣板]報名程序完成",     // document name of 'finihsed' email template, should be located in the same folder
    full: "[email樣板]額滿通知",             // document name of 'full' email template, should be located in the same folder
  },
  tmpl_patt: /from:\s(.*)\ntitle:\s(.*)\nbody:\n((.|\n)*)/,   // pattern of email template
  price_patt: /(\$\d+)/,          // pattern of price
  customer_rec: [6, 1],           // top-left corner of customer records. [row, col]
  
  // info not to be resolved.
  info: {
    "MAX": [4, 2],
    "報名成功人數": [4, 4],
  },

  // info to be resolved.
  resolve: {
    "路線": [2, 3],
    "集合地點內容介紹": [2, 5],
    "日期": [2, 1],
    "開始時間": [2, 2],
    "伴走志工": [2, 4],
  },

  log_sheet: "log",
  is_debug: false,               // won't send mail in debug mode
  is_log: true,                  // logging
};

var Const = {
  sent: 0,              // index to sent flat
  status: 1,            // index to status
  paid_status: 2,       // index to paid status
  email: 5,             // index to email address
  price: 7,             // index to paid price
};


function log_(msg) {
  if (Conf.is_debug == false && Conf.is_log == false) {
    return;
  }
  
  // create a 'Log' sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Conf.log_sheet);
  if (sheet == null) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(Conf.log_sheet);
    if (sheet == null) {
      Logger.log("Unable to create log sheet");
      return;
    }
  }

  sheet.appendRow([msg]);
}


function resolve_(s, sheet, customer) {
  // this function resolve '[...]' in string.
  var ret = s;
  
  for (var key in Conf.resolve) {
    var patt = "(\\[" + key + "\\])";
    var range = sheet.getRange(Conf.resolve[key][0], Conf.resolve[key][1]);
    var val = range.getValue();
   
    // special case for Date object
    if (val instanceof Date) {
      var fmt = range.getNumberFormat();
      if (fmt) {
        // now 'Time' format would produce various unexpected result
        // use 'Plain Text' would make things easier.

        val = Utilities.formatDate(val, "GMT+0800", fmt);
      }
    }
    
    ret = ret.replace(new RegExp(patt, "g"), val);
  }
  
  // resolve 票價
  var patt = /\[票價\]/;
  ret = ret.replace(patt, customer[Const.price]);
  
  return ret;
}


function prepareEmail_(customer, sheet, tmpl) {
  var ret = {
    to: customer[Const.email],
    title: resolve_(tmpl.title, sheet, customer),
    body: resolve_(tmpl.body, sheet, customer),
  };
  
  return ret;
}


function cb_full_(customer, sheet) {
  if (customer[Const.sent] == "V") return false;
  if (customer[Const.email] == "") return false;
  if (customer[Const.status] != "") return false;
  if (customer[Const.paid_status] == "已付全額" || customer[Const.paid_status] == "已付部分") return false; 
  
  // it seems useless to check sheet for each customer...
  // but currently, it's more intuitive to keep code here.
  var the_max = sheet.getRange(Conf.info["MAX"][0], Conf.info["MAX"][1]).getValue();
  var the_total = sheet.getRange(Conf.info["報名成功人數"][0], Conf.info["報名成功人數"][1]).getValue();
  
  if (the_total != null && the_max != null) {
    return the_total >= the_max;
  } else {
    return false;
  }
}

function cb_finished_(customer, _) {
  // skip those customers already got a email
  if (customer[Const.sent] == "V") return false;
  if (customer[Const.email] == "") {
    return false;
  }
  if (customer[Const.status] != "已報名成功") {
    return false;
  }
  
  return true;
}

function handleSheet_(sheet, tmpl, cb) {
  // get datum of all customers
  var customers = sheet.getRange(Conf.customer_rec[0], Conf.customer_rec[1], sheet.getLastRow(), sheet.getLastColumn()).getValues()

  // iterate through all customers
  for (var i = 0; i < customers.length; i++) {
    var curC = customers[i];
      
    if (false == cb(curC, sheet)) continue;

    // trim useless words in paid price of each customer
    var matched = curC[Const.price].match(Conf.price_patt);
    if (matched != null) {
      curC[Const.price] = matched[0];
    } else {
      log_("[error] unable to locate price in string: [" + curC[Const.price] + "]");
        
      // skip this customer.
      continue;
    }

    // prepare email content by resolving those variables.
    var email = prepareEmail_(curC, sheet, tmpl);

    try {
      if (Conf.is_debug == false) {
        MailApp.sendEmail(email.to, email.title, email.body);
          
        // update email sent status
        sheet.getRange(i+Conf.customer_rec[0], 1).setValue("V");
      }
        
      log_("[info] mail sent: [" + email.to + "]");
    } catch (e) {
      log_("[error] unable to send email: [" + e.message + "]");
    }
  }
}


function getEmailTemplate_(name) {
  var tmpl = DriveApp.getFilesByName(name);
  while (tmpl.hasNext()) {
    var f = tmpl.next();
    if (f.getMimeType() != "application/vnd.google-apps.document") continue;

    var doc = DocumentApp.openById(f.getId());
    var tmplText = doc.getBody().getText().match(Conf.tmpl_patt);
    return {
      from: tmplText[1],
      title: tmplText[2],
      body: tmplText[3]
    };
  }

  log_("Unable to load email template[" + name + "]");
  return null;
}


// entry point
function timeDriven(e) {
  log_("begin of timeDriven");
  
  // load email template
  var tmpl = getEmailTemplate_(Conf.tmpl_name.finished)
  if (tmpl != null) {
      handleSheet_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Conf.sheet_name), tmpl, cb_finished_);
  }
  
  log_("end of timeDriven");
}


// entry point
function formSubmit(e) {
  log_("begin of formSubmit");
  
  // load email template
  var tmpl = getEmailTemplate_(Conf.tmpl_name.full);
  if (tmpl != null) {
      handleSheet_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Conf.sheet_name), tmpl, cb_full_);
  }
  
  log_("end of formSubmit");
}
