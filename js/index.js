const BeneReport = require('./js/benereport')

// Common button color definitions amongst all files
const LightBlue = ' style="background-color:#33C3F0;color:#FFF" '
const DarkBlue = ' style="background-color:#3365f0;color:#FFF" '
const ErrorRow = ' style="background-color:#e32636;color:#000" '
const WarnRow = ' style="background-color:#ffbf00;color:#000" '
const OkayRow = ' style="background-color:#5f9ea0;color:#000" '

var ThisWindow

window.onload = function() {

  ThisWindow = window

  showSelectionButtons('')
}

function showSelectionButtons(selectedButton) {

  let updateColor = (selectedButton == 'benes' ? DarkBlue : LightBlue)
  let changePwdColor = (selectedButton == 'tbd1' ? DarkBlue : LightBlue)
  let settingsColor = (selectedButton == 'tbd2' ? DarkBlue : LightBlue)

  // Would use the package version directly but that isn't
  // available in the installed product. So if the package
  // version is defined that we means we're in development
  // so alert the developer if hardcoded version does not
  // match package version.
  let hardVersion = '1.0.1'
  let pkgVersion = process.env.npm_package_version
  if (typeof(pkgVersion) != 'undefined') {
    if (hardVersion != pkgVersion) {
      alert('Hardcoded version should be '+pkgVersion)
    }
  }

  let tableBody = '<tr><p style="text-align:center">'+hardVersion+'</p></tr>'
  tableBody += '<tr><td>'
  tableBody += '<input type="button" class="four columns" value="Beneficiaries Report"' + updateColor + 'onclick="beneficiariesReport()">'
  tableBody += '<input type="button" class="four columns" value="TBD1"' + changePwdColor + 'onclick="TBD(\'tbd1\')">'
  tableBody += '<input type="button" class="four columns" value="TBD2"' + settingsColor + 'onclick="TBD(\'tbd2\')">'
  tableBody += '</td></tr>'

  // Fill the table content
  document.getElementById('table-selections').innerHTML = tableBody
}

function beneficiariesReport() {

  showSelectionButtons('benes')
  BeneReport.start(ThisWindow)
}

function passToBeneReport(str) {
  BeneReport.passTo(str)
}

function TBD(selectedButton) {

  showSelectionButtons(selectedButton)
  document.getElementById('table-main').innerHTML = ''
  document.getElementById('table-main-results').innerHTML = ''
  document.getElementById('table-details').innerHTML = ''
}
