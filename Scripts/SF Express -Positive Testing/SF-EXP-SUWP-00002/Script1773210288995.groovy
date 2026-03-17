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

String filePath = RunConfiguration.getProjectDir() + '/Data Files/SF-EXP-SUWP-00002.xlsx'



WebUI.openBrowser('')
WebUI.maximizeWindow()

WebUI.navigateToUrl('https://sf.ekonek.com/login')

WebUI.setText(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_Username'), GlobalVariable.Username)

WebUI.setEncryptedText(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_Password'), GlobalVariable.Password)

WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_Login'))

WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_Upload XLS'))

// Close the native File Explorer dialog if it remains open after file selection
Robot robot = new Robot()
robot.delay(1000)
robot.keyPress(KeyEvent.VK_ESCAPE)
robot.keyRelease(KeyEvent.VK_ESCAPE)

WebUI.uploadFile(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_uploadxls'), filePath)


WebUI.comment("Excel file uploaded")



// Open the Excel file
FileInputStream fis = new FileInputStream(filePath)
XSSFWorkbook workbook = new XSSFWorkbook(fis)
def sheet = workbook.getSheetAt(0)

DataFormatter formatter = new DataFormatter()

// Reading Row 1 = first data row 
String hawb       = formatter.formatCellValue(sheet.getRow(1).getCell(0))
String eventCode  = formatter.formatCellValue(sheet.getRow(1).getCell(1))
String reasonCode = formatter.formatCellValue(sheet.getRow(1).getCell(2))
String date       = formatter.formatCellValue(sheet.getRow(1).getCell(3))

workbook.close()
fis.close()


//  Check if Excel values are read correctly in the log
WebUI.comment("Expected HAWB: "        + hawb)
WebUI.comment("Expected EVENT_CODE: "  + eventCode)
WebUI.comment("Expected REASON_CODE: " + reasonCode)
WebUI.comment("Expected DATE: "        + date)


// Use data-column-id XPath 

def dynamicXPath = { String xpath ->
	TestObject to = new TestObject()
	to.addProperty('xpath', ConditionType.EQUALS, xpath)
	return to
}

WebUI.waitForElementVisible(dynamicXPath("//div[@id='row-0']//div[@data-column-id='1']"), 10)

String actualHAWB       = WebUI.getText(dynamicXPath("//div[@id='row-0']//div[@data-column-id='1']"))
String actualEventCode  = WebUI.getText(dynamicXPath("//div[@id='row-0']//div[@data-column-id='2']"))
String actualReasonCode = WebUI.getText(dynamicXPath("//div[@id='row-0']//div[@data-column-id='3']"))
String actualDate       = WebUI.getText(dynamicXPath("//div[@id='row-0']//div[@data-column-id='4']"))


WebUI.comment("Actual HAWB: "        + actualHAWB)
WebUI.comment("Actual EVENT_CODE: "  + actualEventCode)
WebUI.comment("Actual REASON_CODE: " + actualReasonCode)
WebUI.comment("Actual DATE: "        + actualDate)

// To verify the uploaded excel value and uploaded in the UI
String actual   = "${actualHAWB} | ${actualEventCode} | ${actualReasonCode} | ${actualDate}"
String expected = "${hawb} | ${eventCode} | ${reasonCode} | ${date}"

// To verify the uploaded excel value and uploaded in the UI
WebUI.verifyMatch(actual, expected, false)

WebUI.comment(" TEST PASSED — 1 row valid data displays correctly")



