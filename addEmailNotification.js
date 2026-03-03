const sheetName = 'Sheet1'
const scriptProp = PropertiesService.getScriptProperties()

function initialSetup() {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  scriptProp.setProperty('key', activeSpreadsheet.getId())
}

function doPost(e) {
  const lock = LockService.getScriptLock()
  lock.tryLock(10000)

  try {
    const doc = SpreadsheetApp.openById(scriptProp.getProperty('key'))
    const sheet = doc.getSheetByName(sheetName)

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    const nextRow = sheet.getLastRow() + 1

    const newRow = headers.map(function(header) {
      return e.parameter[header] || ""
    })

    sheet.getRange(nextRow, 1, 1, newRow.length).setValues([newRow])

    
    const subject = "New Registration - WZC Youth Impact Challenge"
    let body = "New Registration Received:\n\n"

    headers.forEach(function(header, i) {
      body += header + ": " + newRow[i] + "\n"
    })

    MailApp.sendEmail(
      "tejaballa@mobilise.agency",
      subject,
      body
    )

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON)

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON)

  } finally {
    lock.releaseLock()
  }
}

