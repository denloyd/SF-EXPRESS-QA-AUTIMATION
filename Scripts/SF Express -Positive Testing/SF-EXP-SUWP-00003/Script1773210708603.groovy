import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testng.keyword.TestNGBuiltinKeywords as TestNGKW
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable
import org.openqa.selenium.Keys as Keys
import com.kms.katalon.core.testobject.ConditionType
import com.kms.katalon.core.testobject.TestObject
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.ss.usermodel.DataFormatter
import java.io.FileInputStream
import java.awt.Robot
import java.awt.event.KeyEvent

import com.kms.katalon.core.configuration.RunConfiguration

String filePath = RunConfiguration.getProjectDir() + '/Data Files/SF-EXP-SUWP-00003.xlsx'


// Read ALL rows from Excel BEFORE uploading

FileInputStream fis = new FileInputStream(filePath)
XSSFWorkbook workbook = new XSSFWorkbook(fis)
def sheet = workbook.getSheetAt(0)
DataFormatter formatter = new DataFormatter()

// Store all data rows in a list
List<Map<String, String>> excelRows = []
for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
    def row = sheet.getRow(rowNum)
    if (row == null) continue
    excelRows.add([
        hawb       : formatter.formatCellValue(row.getCell(0)),
        eventCode  : formatter.formatCellValue(row.getCell(1)),
        reasonCode : formatter.formatCellValue(row.getCell(2)),
        date       : formatter.formatCellValue(row.getCell(3))
    ])
}

workbook.close()
fis.close()

WebUI.comment("Total rows read from Excel: " + excelRows.size())
excelRows.eachWithIndex { row, i ->
    WebUI.comment("Row ${i + 1} — HAWB: ${row.hawb} | EVENT: ${row.eventCode} | REASON: ${row.reasonCode} | DATE: ${row.date}")
}



// Login and upload the Excel file
WebUI.openBrowser('')
WebUI.maximizeWindow()
WebUI.navigateToUrl('https://sf.ekonek.com/login')

WebUI.setText(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_Username'), GlobalVariable.Username)

WebUI.setEncryptedText(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_Password'), GlobalVariable.Password)

WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_Login'))

WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_Upload XLS'))

Robot robot = new Robot()
robot.delay(1000)
robot.keyPress(KeyEvent.VK_ESCAPE)
robot.keyRelease(KeyEvent.VK_ESCAPE)

WebUI.uploadFile(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_uploadxls'), filePath)
WebUI.comment("Excel file uploaded — ${excelRows.size()} rows")



// Verify EACH row in UI matches Excel
def dynamicXPath = { String xpath ->
    TestObject to = new TestObject()
    to.addProperty('xpath', ConditionType.EQUALS, xpath)
    return to
}

WebUI.waitForElementVisible(dynamicXPath("//div[@id='row-0']//div[@data-column-id='1']"), 10)


for (int i = 0; i < excelRows.size(); i++) {
    Map<String, String> excelRow = excelRows[i]

    // Get actual values from UI
    String actualHAWB       = WebUI.getText(dynamicXPath("//div[@id='row-${i}']//div[@data-column-id='1']"))
    String actualEventCode  = WebUI.getText(dynamicXPath("//div[@id='row-${i}']//div[@data-column-id='2']"))
    String actualReasonCode = WebUI.getText(dynamicXPath("//div[@id='row-${i}']//div[@data-column-id='3']"))
    String actualDate       = WebUI.getText(dynamicXPath("//div[@id='row-${i}']//div[@data-column-id='4']"))

    // Print expected vs actual for this row
   WebUI.comment(" Row ${i + 1} ")
   WebUI.comment("Expected HAWB (Excel): ${excelRow.hawb} | Actual HAWB (UI): ${actualHAWB}")
   WebUI.comment("Expected EVENT_CODE (Excel): ${excelRow.eventCode} | Actual EVENT_CODE (UI): ${actualEventCode}")
   WebUI.comment("Expected REASON_CODE (Excel): ${excelRow.reasonCode} | Actual REASON_CODE (UI): ${actualReasonCode}")
   WebUI.comment("Expected DATE (Excel): ${excelRow.date} | Actual DATE (UI): ${actualDate}")

    String actual   = "${actualHAWB} | ${actualEventCode} | ${actualReasonCode} | ${actualDate}"
	String expected = "${excelRow.hawb} | ${excelRow.eventCode} | ${excelRow.reasonCode} | ${excelRow.date}"
	
	WebUI.verifyMatch( actual, expected, false)

}


WebUI.comment(" TEST PASSED — All ${excelRows.size()} rows verified successfully")