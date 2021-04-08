const XLSX = require('xlsx')
const fs = require('fs')
var path = require('path')

// Sheet names
const AvailBalSheetName = 'Available Balance'
const CurrentBeneSheetName = 'Current Beneficiaries'
const ExpiredSheetName = 'Expired'
const DeceasedSheetName = 'Deceased'
const ReleasedSheetName = 'Released'
const UsedUpSheetName = 'Used Up'

// Column names and contents
const CardStatusColName = 'Card Status'
const CardStatusClosed = 'Closed'
const CardStatusBlocked = 'Blocked'
const FirstNameColName = 'First Name'
const LastNameColName = 'Last Name'
const AdminFirstName = 'SJWHTF Admin'
const AdminLastName = 'Assistant'
const ExpDateColName = 'Expiration Date'
const AvailBalColName = 'Available Balance'
const LedgerBalColName = 'Ledger Balance'
const LastUsedDateColName = 'Last Used Date'
const EmailColName = 'Email'
const CardNumberColName = 'Card Number'
const GroupNameColName = 'Group Name'
const SpendingRuleSetColName = 'Spending Ruleset Name'
const AccountIdColName = 'Account ID'
const CustomIdColName = 'CustomId'
const GrantExpColName = 'Grant Exp'
const GrantNumColName = 'Grant #'
const DateCardClosedColName = 'Date PEX Card Closed'

// Columns to add
const RevExpDateColName = 'Revised Exp Date'

// Common button color definitions amongst all files
const LightBlue = ' style="background-color:#33C3F0;color:#FFF" '
const DarkBlue = ' style="background-color:#3365f0;color:#FFF" '

// Global variables
var MainWindow
var Today
var LogFileName = ''
var LogBuffer = ''
var Details = []

exports.start = function(win) {

  MainWindow = win
  showBeneficiariesReportForm('','')
}

exports.passTo = function(str) {

  if (str == 'handleBeneFileOpen') {
    handleBeneFileOpen()
  } else if (str == 'handleExpDateFileOpen') {
    handleExpDateFileOpen()
  } else if (str == 'processBeneFile') {
    processBeneFile()
  } else if (str == 'showDetails') {
    showDetails()
  } else if (str == 'hideDetails') {
    hideDetails()
  }
}

// Populates the main command table
function showBeneficiariesReportForm(beneFileName,expDateFileName) {

  clearResults()
  clearDetails()

  if (beneFileName == '') {
    beneFileName = getInputFileName('beneFile','')
  }
  if (expDateFileName == '') {
    expDateFileName = getInputFileName('expDateFile','')
  }

  // Generate the table body
  let tableBody = '<tr><td>'
  tableBody += '<input type="button" class="two columns" value="Open"' + LightBlue + 'onclick="passToBeneReport(\'handleBeneFileOpen\')">'
  tableBody += '<input type="text" class="ten columns" value="'+beneFileName+'" placeholder="Click Open to select Beneficiaries Report file" id="beneFile">'
  tableBody += '</td></tr>'
  tableBody += '<tr><td>'
  tableBody += '<input type="button" class="two columns" value="Open"' + LightBlue + 'onclick="passToBeneReport(\'handleExpDateFileOpen\')">'
  tableBody += '<input type="text" class="ten columns" value="'+expDateFileName+'" placeholder="Click Open to select Grant Expiration Date file" id="expDateFile">'
  tableBody += '</td></tr>'
  tableBody += '<tr><td><input type="button" class="three columns" value="Process File"' + LightBlue + 'onclick="passToBeneReport(\'processBeneFile\')"></td></tr>'


  // Fill the table content
  document.getElementById('table-main').innerHTML = tableBody

  document.getElementById('beneFile').focus()
}

function handleBeneFileOpen() {

  clearResults()
  clearDetails()

  // Use system dialog to select file name
  const { dialog } = require('electron').remote
  promise = dialog.showOpenDialog()
  promise.then(
    result => handleBeneFileOpenResult(result['filePaths'][0]),
    error => alert(error)
  )
}

function handleBeneFileOpenResult(fileName) {

  if (typeof(fileName) == 'undefined') {
    showResults(['No Beneficiaries Report file selected.'])
    return
  }

  let csvSuffix = fileName.toLowerCase().indexOf('.csv')
  if (csvSuffix == -1) {
    showResults(['The Beneficiaries Report must be a csv file.'])
    return
  }

  // Generate log file name and today's date
  LogFileName = fileName.substring(0,csvSuffix)+' Activity Log.txt'
  Today = new Date()
  Today.setHours(0, 0, 0, 0)

  // Show file names in the text box
  showBeneficiariesReportForm(fileName,'')
  showResults(['Click "Process File" to create Beneficiaries Form workbook.'])
}

function handleExpDateFileOpen() {

    clearResults()
    clearDetails()

    // Use system dialog to select file name
    const { dialog } = require('electron').remote
    promise = dialog.showOpenDialog()
    promise.then(
      result => handleExpDateFileOpenResult(result['filePaths'][0]),
      error => alert(error)
    )
}

function handleExpDateFileOpenResult(fileName) {

  if (typeof(fileName) == 'undefined') {
    showResults(['No Grant Expiration Date file selected.'])
    return
  }

  let xlsxSuffix = fileName.toLowerCase().indexOf('.xlsx')
  if (xlsxSuffix == -1) {
    showResults(['The Grant Expiration Date must be an xlsx file.'])
    return
  }

  // Show file names in the text box
  showBeneficiariesReportForm('',fileName)
  showResults(['Click "Process File" to create Beneficiaries Form workbook.'])
}

function getInputFileName(elemId,errString) {

  let handle = document.getElementById(elemId)
  if (handle == null) {
    showResults([errString])
    return ''
  }
  return handle.value
}

function isFileType(fileName, type) {

  let suffix = fileName.toLowerCase().indexOf('.'+type)
  if (suffix == -1) {
    return false
  }
  return true
}

function processBeneFile() {

  clearDetails()

  let beneFileName = getInputFileName('beneFile','Click Open to select Beneficiaries Report file')
  if (beneFileName == '') {
    showResults(['No Beneficiaries Report file specified.'])
    return
  }

  // Check file types again since they could have typed it in
  let typePos = beneFileName.toLowerCase().indexOf('.csv')
  if (typePos == -1) {
    showResults(['Beneficiaries Report must be a csv file'])
    return
  }

  showResults(['Working...'])
  clearDetails()

  // Open beneficiaries report file
  let sheet = []
  try {
    let workbook = XLSX.readFile(beneFileName)

    // A csv file only has one sheet
    let sheetName = workbook.SheetNames[0]
    let xlsxSheet = workbook.Sheets[sheetName]
    sheet = XLSX.utils.sheet_to_json(xlsxSheet,{defval:''})
  } catch (e) {
    showResults([e.message])
    return
  }

  // Save original beneficiaries report column order
  let keys
  try {
    keys = Object.keys(sheet[0])
  } catch (e) {
    showResults(['The format of "'+beneFileName+'" is invalid.','Please verify that it is a csv file.'])
    return
  }

  // Set up correct output column order
  for (let i = keys.length-1; i >= 0; i--) {
    let key = keys[i]

    // Remove unneeded beneficiaries report columns
    if (key==EmailColName || key==CardNumberColName || key==LedgerBalColName || key==GroupNameColName || key==SpendingRuleSetColName || key==AccountIdColName || key==CustomIdColName) {
      keys.splice(i,1)
    }

    // Add new column for COVID-19
    if (key==AvailBalColName) {
      keys.splice(i,0,RevExpDateColName)
    }
  }

  // Get expiration dates from Excel workbook
  let expirationDatesByName = {}
  let grantNumbersByName = {}
  let expDatesWorkbook
  try {
    expDatesWorkbook = getGrantExpirationWorkbook()
    expirationDatesByName = getGrantExpirationDates(expDatesWorkbook, grantNumbersByName)
  } catch (e) {
    showResults([e.message])
    return
  }

  // Fix or remove cells in rows from existing sheet
  for (let i = 0; i < sheet.length; i++) {
    let row = sheet[i]
    let name = row[FirstNameColName].toLowerCase()+' '+row[LastNameColName].toLowerCase()

    // Skip cards that will be discarded below
    if (row[CardStatusColName] == CardStatusClosed) {
      continue
    }

    // Copy/clear expiration date column
    if (name in expirationDatesByName) {
      row[ExpDateColName] = expirationDatesByName[name]
    } else {
      row[ExpDateColName] = ''
    }

    // Add new column for COVID-19
    let expDate = new Date(row[ExpDateColName])
    let expYear = expDate.getFullYear()
    if (expYear == '2020' || expYear == '2021') {
      row[RevExpDateColName] = '12/31/2021'
    } else {
      row[RevExpDateColName] = ''
    }

    // Append Grant # to last name
    let grantNum = grantNumbersByName[name]
    if (typeof(grantNum) != 'undefined') {
      row[LastNameColName] = row[LastNameColName] + ' (' + grantNum + ')'
    }

    // Deal with ledger balance and available balance differences
    if (row[LedgerBalColName] != row[AvailBalColName]) {
      //row[AvailBalColName] = row[LedgerBalColName]
      let str = 'Please get PEX pending date and copy to "Last Used Date" for '+row[FirstNameColName]+' '+row[LastNameColName]
      addDetails(str)
      writeLog(str)
    }

    // Deal with used up cards
    if (row[AvailBalColName] == 0 && row[LastUsedDateColName] != '') {
      let str = 'Please follow "Used Up" instructions for '+row[FirstNameColName]+' '+row[LastNameColName]
      addDetails(str)
      writeLog(str)
    }

    // Remove unneeded columns from row
    delete row[EmailColName]
    delete row[CardNumberColName]
    delete row[GroupNameColName]
    delete row[SpendingRuleSetColName]
    delete row[AccountIdColName]
    delete row[CustomIdColName]
    delete row[LedgerBalColName]
  }

  // Create new array without closed cards and admin row
  let closedCount = adminCount = 0
  let newSheet = []
  for (let i = 0; i < sheet.length; i++) {
    let row = sheet[i]
    let name = row[FirstNameColName]+' '+row[LastNameColName]
    if (row[CardStatusColName] == CardStatusClosed) {
      closedCount++
      writeLog('dropped: ' + name)
    } else if (row[FirstNameColName] == AdminFirstName && row[LastNameColName] == AdminLastName) {
      adminCount++
      writeLog('dropped admin: ' + name)
    } else {
      if (row[CardStatusColName] == CardStatusBlocked) {
        newSheet.unshift(row)
      } else {
        newSheet.push(row)
      }
    }
  }

  // Generate output file name using input file name
  let outputFileName = beneFileName.substring(0, typePos)+'.xlsx'

  // If output file already exists, don't overwrite it
  try {
    if (fs.existsSync(outputFileName)) {
      showResults(['Output file "'+outputFileName+'" exists. Please delete it.'])
      return
    }
  } catch(err) {
    // Normal path
  }

  let sts = 'Found '+sheet.length+' cards. Copied '+newSheet.length+' cards. Dropped '+closedCount+' closed cards and '+adminCount+' admin cards.'

  try {
    // Set up output workbook
    let newWorkbook = XLSX.utils.book_new()
    newWorkbook.SheetNames.push(AvailBalSheetName)
    let newWorkSheet = XLSX.utils.json_to_sheet(newSheet,{header:keys})
    newWorkbook.Sheets[AvailBalSheetName] = newWorkSheet

    // Add sheets from Expiration Dates workbook
    addSheet(newWorkbook, expDatesWorkbook, ExpiredSheetName)
    addSheet(newWorkbook, expDatesWorkbook, DeceasedSheetName)
    addSheet(newWorkbook, expDatesWorkbook, ReleasedSheetName)
    addSheet(newWorkbook, expDatesWorkbook, UsedUpSheetName)

    // Write output workbook
    XLSX.writeFile(newWorkbook, outputFileName)

    // Write log file
    writeLog(sts)
    syncLog()
  } catch (e) {
      showResults([e.message])
      return
  }

  showResults([sts])
  revealDetails('Further actions required. See details.')
}

// Populates the results table
function showResults(status) {

  // Generate the table body
  let tableBody = ''
  for (let i = 0; i < status.length; i++) {
    tableBody += '<tr>' + status[i] + '</tr>'
  }

  // Fill the table content
  document.getElementById('table-main-results').innerHTML = tableBody
}

// clears the results table
function clearResults() {

    // clear the table content
    document.getElementById('table-main-results').innerHTML = ''
}

function addDetails(str) {

  Details.push('<tr><td>'+str+'</td></tr>')
}

function clearDetails(str) {

  Details = []
  document.getElementById('table-details').innerHTML = ''
}

function showDetails() {

  // Generate the table body
  let tableBody = '<td><input type="button" value="Hide Details"' + LightBlue + 'onclick="passToBeneReport(\'hideDetails\')"></td>'
  for (let i = 0; i < Details.length; i++) {
    tableBody += Details[i]
  }

  // Fill the table content
  document.getElementById('table-details').innerHTML = tableBody
}

function revealDetails(str) {

  if (Details.length > 0) {
    let tableBody = document.getElementById('table-main-results').innerHTML
    tableBody += '<tr>'+str+'</tr>'
    document.getElementById('table-main-results').innerHTML = tableBody
    hideDetails()
  }
}

function hideDetails() {

  let tableBody = '<td><input type="button" value="Show Details"' + LightBlue + 'onclick="passToBeneReport(\'showDetails\')"></td>'
  document.getElementById('table-details').innerHTML = tableBody
}

function writeLog(str) {

  console.log(str)
  if (LogBuffer == '') {
    let month = Today.getMonth()+1
    let day = Today.getUTCDay()
    let year = Today.getFullYear()
    LogBuffer += 'Generated report on '+month+'/'+day+'/'+year+'\n'
  }

  LogBuffer += str +'\n'
}

function syncLog() {

  fs.writeFileSync(LogFileName, LogBuffer, 'utf-8')
}

function getGrantExpirationWorkbook() {

  let expDateFileName = getInputFileName('expDateFile','')
  if (expDateFileName == '') {
    throw {
      name:'File Error',
      message:'Grant Expiration Date file is not specified.'
    }
  }
  if (isFileType(expDateFileName,'xlsx') == false) {
    throw {
      name:'File Error',
      message:'Grant Expiration Date must be an xlsx file.'
    }
  }

  return XLSX.readFile(expDateFileName)
}

function getGrantExpirationDates(workbook,grantNumbersByName) {

  let xlsxSheet = getSheet(workbook,CurrentBeneSheetName)
  let expDateSheet = XLSX.utils.sheet_to_json(xlsxSheet,{raw:false})
  let expDatesByName = {}
  for (let i = 0; i < expDateSheet.length; i++) {
    let row = expDateSheet[i]
    let name = row[FirstNameColName].toLowerCase()+' '+row[LastNameColName].toLowerCase()

    // Use 4-digit year assuming mm/dd/yy format
    let date = row[GrantExpColName]
    let slash = date.lastIndexOf('/')
    if (slash > -1) {
      let end = date.length
      date = date.substring(0,slash+1)+'20'+date.substring(slash+1,end)
    }

    expDatesByName[name] = date

    grantNumbersByName[name] = row[GrantNumColName]
  }

  return expDatesByName
}

function getSheet(workbook, sheetName) {

  let sheet = workbook.Sheets[sheetName]
  if (typeof(sheet) == 'undefined') {
    throw {
      name:'Sheet Error',
      message:'Grant Expiration Date file has no "'+sheetName+'" sheet.'
    }
  }
  return sheet
}

function addSheet(dstWorkbook, srcWorkbook, sheetName) {

  const thirtyDays = 30*24*60*60*1000
  let todayInMillis = Today.getTime()

  let emptyRow = {}
  emptyRow[FirstNameColName] = ''
  emptyRow[LastNameColName] =''
  emptyRow[GrantExpColName] = ''
  emptyRow[DateCardClosedColName] = ''

  // Get sheet from source workbook
  let xlsxSheet = getSheet(srcWorkbook,sheetName)
  let sheet = XLSX.utils.sheet_to_json(xlsxSheet,{raw:false})

  // only keep cards with closed date equal to less than 30 days ago
  let newSheet = []
  for (let i = sheet.length-1; i >= 0; i--) {
    let row = sheet[i]
    let str = row[DateCardClosedColName]
    let dateInMillis = Date.parse(str)

    if (dateInMillis >= (todayInMillis - thirtyDays)) {
      // Add Grant # to last name
      let grantNum = row[GrantNumColName]
      if (typeof(grantNum) != 'undefined') {
        row[LastNameColName] = row[LastNameColName] + ' (' + grantNum + ')'
      }

      newSheet.push(row)
    }
  }

  // If no rows were copied, insert an empty one so we get a header row
  if (newSheet.length == 0) {
    newSheet.push(emptyRow)
  }

  // Add sheet to destination workbook
  dstWorkbook.SheetNames.push(sheetName)
  xlsxSheet = XLSX.utils.json_to_sheet(newSheet)
  dstWorkbook.Sheets[sheetName] = xlsxSheet
}
