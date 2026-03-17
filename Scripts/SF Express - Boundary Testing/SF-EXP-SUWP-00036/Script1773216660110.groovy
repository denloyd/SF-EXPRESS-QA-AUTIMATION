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

String filePath = RunConfiguration.getProjectDir() + '/Data Files/SF-EXP-SUWP-00036.xlsx'

// Read ALL DATE values from Excel BEFORE uploading

FileInputStream fis = new FileInputStream(filePath)
XSSFWorkbook workbook = new XSSFWorkbook(fis)
def sheet = workbook.getSheetAt(0)
DataFormatter formatter = new DataFormatter()

List<String> excelDates = []
for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
	def row = sheet.getRow(rowNum)
	if (row == null) continue

	// Get date cell
	def dateCell = row.getCell(3)

	// SKIP row if DATE cell is null or empty
	if (dateCell == null) continue
	String dateValue = formatter.formatCellValue(dateCell).trim()
	if (dateValue.isEmpty()) continue

	excelDates.add(dateValue)
}

workbook.close()
fis.close()

WebUI.comment("Total valid rows read from Excel: " + excelDates.size())
excelDates.eachWithIndex { d, i ->
	WebUI.comment("Row ${i + 1} DATE (Excel): [" + d + "]")
}


// Validate ALL Excel DATE formats are mm/dd/yyyy hh:mm:ss

String datePattern = /^\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}:\d{2}$/

excelDates.eachWithIndex { d, i ->
	if (d ==~ datePattern) {
		WebUI.comment(" Row ${i + 1} Excel DATE format VALID: [" + d + "]")
	} else {
		WebUI.comment(" Row ${i + 1} Excel DATE format INVALID: [" + d + "]")
		assert false, "Row ${i + 1} DATE in Excel does not match mm/dd/yyyy hh:mm:ss — Actual: [" + d + "]"
	}
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
WebUI.comment(" Excel file uploaded — " + excelDates.size() + " rows")


// Verify DATE format in UI for ALL rows

def dynamicXPath = { String xpath ->
	TestObject to = new TestObject()
	to.addProperty('xpath', ConditionType.EQUALS, xpath)
	return to
}

WebUI.waitForElementVisible(dynamicXPath("//div[@id='row-0']//div[@data-column-id='4']"), 10)

for (int i = 0; i < excelDates.size(); i++) {
	String actualDate = WebUI.getText(dynamicXPath("//div[@id='row-${i}']//div[@data-column-id='4']"))

	WebUI.comment(" Row ${i + 1} ")
	WebUI.comment("Expected DATE (Excel): [" + excelDates[i] + "]")
	WebUI.comment("Actual DATE (UI): ["      + actualDate    + "]")

	if (actualDate ==~ datePattern) {
		WebUI.comment("Row ${i + 1} UI DATE format VALID: [" + actualDate + "]")
	} else {
		WebUI.comment(" Row ${i + 1} UI DATE format INVALID: [" + actualDate + "]")
		assert false, "Row ${i + 1} DATE in UI does not match mm/dd/yyyy hh:mm:ss — Actual: [" + actualDate + "]"
	}
}

WebUI.comment(" TEST PASSED — All " + excelDates.size() + " row(s) have valid DATE format mm/dd/yyyy hh:mm:ss")