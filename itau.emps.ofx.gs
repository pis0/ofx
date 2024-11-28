const CONFIG = {
  menu: {
    menulabel: "Utils",
    itemLabel: "Export as OFX file"
  },
  file: {
    namePrefix: "itau.emps",
    type: "application/x-msmoney",
    header: {
      ofxHeader:100,
      data: "OFXSGML", 
      version: 102, 
      security: "NONE", 
      encondig: "USASCII",
      charset: 1252, 
      compression: "NONE", 
      oldFileUid: "NONE", 
      newFileUid: "NONE"
    },
    top: {
      language: "POR",
      curdef: "BRL",
      trnuid: 1001,
      acctType: "CHECKING",
      status: {
        code: 0,
        severity: "INFO",
      }
    }
  },
  dialog: {
    title: "OFX File Created",
    btnLabel: "VIEW"
  },
  sheetLayoutRef: {
    sheetName: "Lançamentos",
    dtStartEndSeparator: "até",
    dateColRef: "Data",
    moneyColRef: "Valor (R$)", 
    releasesColRef: "Lançamento",
    pjorpfColRef: "CNPJ/CPF",
    nameColRef: "Razão social"
  },
  aux: {
    timeFuse: "[-03:EST]"
  }
};

function onOpen() { 
  SpreadsheetApp.getUi()
    .createMenu(CONFIG.menu.menulabel)
    .addItem(CONFIG.menu.itemLabel, 'downloadOFX') 
    .addToUi();
} 

function downloadOFX() {
  const ofx = ofxFile();
  const fileName = generateFileName();  
  
  try {
    const file = getOrCreateFile(fileName, ofx);  
    showDownloadDialog(file);
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Error: ${error.message}`);
  }
}

function generateFileName() {
  return `${CONFIG.file.namePrefix}_${new Date().toISOString().split('T')[0]}.ofx`;
}

function getOrCreateFile(fileName, content) {
  const chkFile = DriveApp.getFilesByName(fileName);
  let file; 
  
  if (chkFile.hasNext()) {
    file = chkFile.next();
    file.setContent(content);  // Update content if the file exists
  } else {
    file = DriveApp.createFile(fileName, content, CONFIG.file.type);  // Create a new file if it doesn't exist
  }
  
  return file;
}

function showDownloadDialog(file) {
  const anchor = `<a href="${file.getUrl()}" target="_blank"><button>${CONFIG.dialog.btnLabel}</button></a>`; 
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(anchor), CONFIG.dialog.title);
}

function getSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheetLayoutRef.sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${CONFIG.sheetLayoutRef.sheetName}" not found.`);
  }
  return sheet;
}

function formatDate(date) {
  return date.toISOString() 
  .substring(0, 19)
  .replace(/T/g, "")
  .replace(/-/g, "")
  .replace(/:/g, ""); 
}

function header() {
  return `
OFXHEADER:${CONFIG.file.header.ofxHeader} 
DATA:${CONFIG.file.header.data}  
VERSION:${CONFIG.file.header.version}  
SECURITY:${CONFIG.file.header.security}  
ENCODING:${CONFIG.file.header.encondig}  
CHARSET:${CONFIG.file.header.charset}  
COMPRESSION:${CONFIG.file.header.compression}  
OLDFILEUID:${CONFIG.file.header.oldFileUid}  
NEWFILEUID:${CONFIG.file.header.newFileUid} 
  `;
}

function parseDateRange() {
  const sheet = getSheet();
  const rawDates = sheet.getRange('B8:C8').getValues().flat().filter(Boolean).join('').replace(/ /g, "").split(CONFIG.sheetLayoutRef.dtStartEndSeparator);
  
  if (rawDates.length !== 2) {
    throw new Error(`Invalid date range format in B8:C8. Expected format: dd/mm/yyyy ${CONFIG.sheetLayoutRef.dtStartEndSeparator}"" dd/mm/yyyy.`);
  }
  
  const dtStart = formatDate(new Date(rawDates[0].split('/').reverse().join('/')));
  const dtEnd = formatDate(new Date(rawDates[1].split('/').reverse().join('/')));
  
  return [dtStart, dtEnd];
}

function top() {
  const dtServer = formatDate(new Date());
  const bankId = "0341";  // Bank code for Itaú
  const agcc = getSheet().getRange('B4:B5').getValues().flat().filter(Boolean).join("").replace(/-/g, "");

  const [dtStart, dtEnd] = parseDateRange();
  
  return `
<OFX> 
<SIGNONMSGSRSV1>  
<SONRS> 
<STATUS> 
<CODE>${CONFIG.file.top.status.code}</CODE> 
<SEVERITY>${CONFIG.file.top.status.severity}</SEVERITY> 
</STATUS> 
<DTSERVER>${dtServer}${CONFIG.aux.timeFuse}</DTSERVER> 
<LANGUAGE>${CONFIG.file.top.language}</LANGUAGE> 
</SONRS> 
</SIGNONMSGSRSV1> 
<BANKMSGSRSV1> 
<STMTTRNRS> 
<TRNUID>${CONFIG.file.top.trnuid}</TRNUID> 
<STATUS> 
<CODE>${CONFIG.file.top.status.code}</CODE> 
<SEVERITY>${CONFIG.file.top.status.severity}</SEVERITY> 
</STATUS> 
<STMTRS> 
<CURDEF>${CONFIG.file.top.curdef}</CURDEF>  
<BANKACCTFROM> 
<BANKID>${bankId}</BANKID> 
<ACCTID>${agcc}</ACCTID> 
<ACCTTYPE>${CONFIG.file.top.acctType}</ACCTTYPE>  
</BANKACCTFROM> 
<BANKTRANLIST> 
<DTSTART>${dtStart}${CONFIG.aux.timeFuse}</DTSTART> 
<DTEND>${dtEnd}${CONFIG.aux.timeFuse}</DTEND>
`;
}

function bottom() {

  let date = null
  try {
    const rawDateAndHour = getSheet().getRange('B2:C2').getValues().flat().filter(Boolean).join(" ");
    const rawDateAndHourSplit = rawDateAndHour.split(" "); 
    const timeSplit = rawDateAndHourSplit[1].split(":");
    date = new Date(rawDateAndHourSplit[0].split('/').reverse().join('/'));
    date.setHours(parseInt(timeSplit[0]));
    date.setMinutes(parseInt(timeSplit[1])); 
    date.setSeconds(parseInt(timeSplit[2]));
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Error: ${error.message}`); 
  }
  
  const dtAsof = formatDate(date ? date : new Date());
  
  return `
</BANKTRANLIST> 
<LEDGERBAL> 
<BALAMT>${balamt}</BALAMT> 
<DTASOF>${dtAsof}${CONFIG.aux.timeFuse}</DTASOF> 
</LEDGERBAL> 
</STMTRS> 
</STMTTRNRS> 
</BANKMSGSRSV1> 
</OFX>`;
}

let balamt = 0;

function releases() {
  let releasesXML = "";
  const sheet = getSheet();
  const values = sheet.getRange('A10:F' + sheet.getLastRow()).getValues();
  const firstRow = values[0];
  
  const dateColRef = firstRow.indexOf(CONFIG.sheetLayoutRef.dateColRef);
  const moneyColRef = firstRow.indexOf(CONFIG.sheetLayoutRef.moneyColRef);
  const launchesColRef = firstRow.indexOf(CONFIG.sheetLayoutRef.releasesColRef);
  const pjorpfColRef = firstRow.indexOf(CONFIG.sheetLayoutRef.pjorpfColRef);
  const nameColRef = firstRow.indexOf(CONFIG.sheetLayoutRef.nameColRef);

  if (dateColRef === -1 || moneyColRef === -1 || launchesColRef === -1) {
    throw new Error('Missing required columns in the data. Please check your sheet headers.');
  }

  let lastDate = "";
  let releasesCountAtSameDay = 1;
  
  for (let i = 1; i < values.length; i++) {
    const money = values[i][moneyColRef];
    if (!money) continue; 

    balamt += parseFloat(money);
    
    const rawDate = values[i][dateColRef];
    const type = parseFloat(money) < 0 ? "DEBIT" : "CREDIT";
    const description = `${values[i][launchesColRef].trim()} ${values[i][pjorpfColRef].trim()} ${values[i][nameColRef].trim()}`;
    
    const dateValues = rawDate.split("/");
    const date = formatDate(new Date(dateValues[2], Number(dateValues[1]) - 1, dateValues[0])).substring(0, 8);

    if (lastDate !== date) {
      releasesCountAtSameDay = 1;
      lastDate = date;
    } else {
      releasesCountAtSameDay++;
    }

    const zeroPad = (num, places) => String(num).padStart(places, '0');
    const numAux = zeroPad(releasesCountAtSameDay, 2);

    releasesXML += `
<STMTTRN> 
<TRNTYPE>${type}</TRNTYPE> 
<DTPOSTED>${date}</DTPOSTED> 
<TRNAMT>${money}</TRNAMT> 
<FITID>${date}${numAux}</FITID> 
<CHECKNUM>${date}${numAux}</CHECKNUM> 
<MEMO>${description}</MEMO> 
</STMTTRN>
`;
  }

  return releasesXML;
}

function ofxFile() {
  let ofx = header();
  ofx += top();
  ofx += releases();
  ofx += bottom();
  return ofx;
}
