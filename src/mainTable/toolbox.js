function gif(text) {
  let gifUrl = 'https://cdn-images-1.medium.com/max/1600/1*9EBHIOzhE1XfMYoKz1JcsQ.gif';

  if (!text) text = ' ';
  let gifTemplate = '<img src="' + gifUrl + '" alt="Progress Indicator" style="width: 100%;height: auto; ">';
  let html = HtmlService.createHtmlOutput(gifTemplate).setHeight(225).setWidth(300);
  SpreadsheetApp.getUi().showModalDialog(html, text);
};

function alertDialog(text) {
  SpreadsheetApp.getUi().alert(text);
};

function closeForm(text) {
  if (!text) text = ' ';
  let html = HtmlService.createTemplateFromFile('formReadyClose.html').evaluate().setHeight(20).setWidth(50);
  SpreadsheetApp.getUi().showModalDialog(html, text);
};

function getStrValue(value) {
  return ('' + value).trim();
}

function getNumValue(value) {
  return (+ value);
}

function getDateValue(value) {
  return new Date(value);
}

function getDateValueAsString(value, fStr = "%d.%m.%Y") {
  if (!value || !(value instanceof Date))
    return value;

  return dateFormat(value, fStr);
}

function getBoolValue(value) {
  return Boolean(value);
}

function isIterable(obj) {
  if (obj == null) {
    return false;
  }
  return typeof obj[Symbol.iterator] === 'function';
}

function dateFormat(date, fStr = "%Y-%m-%d %H:%M:%S", utc = false) {
  utc = utc ? 'getUTC' : 'get';
  return fStr.replace(/%[YmdHMS]/g, function (m) {
    switch (m) {
      case '%Y': return date[utc + 'FullYear'](); // no leading zeros required
      case '%m': m = 1 + date[utc + 'Month'](); break;
      case '%d': m = date[utc + 'Date'](); break;
      case '%H': m = date[utc + 'Hours'](); break;
      case '%M': m = date[utc + 'Minutes'](); break;
      case '%S': m = date[utc + 'Seconds'](); break;
      default: return m.slice(1); // unknown code, remove %
    }
    // add leading zero if required
    return ('0' + m).slice(-2);
  });
}