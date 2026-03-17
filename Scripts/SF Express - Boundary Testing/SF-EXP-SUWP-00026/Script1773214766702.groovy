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
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.ss.usermodel.DataFormatter
import java.io.FileInputStream
import java.awt.Robot
import java.awt.event.KeyEvent

import com.kms.katalon.core.configuration.RunConfiguration

String filePath = RunConfiguration.getProjectDir() + '/Data Files/SF-EXP-SUWP-00026.xlsx'



// Read ALL HAWB values from Excel BEFORE uploading
FileInputStream fis = new FileInputStream(filePath)
XSSFWorkbook workbook = new XSSFWorkbook(fis)
def sheet = workbook.getSheetAt(0)
DataFormatter formatter = new DataFormatter()

// Store all HAWB values
List<Map<String, String>> excelRows = []
for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
	def row = sheet.getRow(rowNum)
	if (row == null) continue

	def hawbCell = row.getCell(0)
	if (hawbCell == null) continue
	String hawbValue = formatter.formatCellValue(hawbCell).trim()
	if (hawbValue.isEmpty()) continue

	excelRows.add([
		hawb    : hawbValue,
		rowNum  : String.valueOf(rowNum + 1)  
	])
}

workbook.close()
fis.close()

WebUI.comment("Total rows read from Excel: " + excelRows.size())


// Check ALL HAWB values against max length
int MAX_HAWB_LENGTH = 20  
boolean hasInvalidHAWB = false

excelRows.eachWithIndex { row, i ->
	int hawbLength = row.hawb.length()
	WebUI.comment("Row ${i + 1} — HAWB: [" + row.hawb + "] | Length: " + hawbLength + " | Max: " + MAX_HAWB_LENGTH)

	if (hawbLength > MAX_HAWB_LENGTH) {
		WebUI.comment(" Row ${i + 1} HAWB EXCEEDS max length — [" + row.hawb + "] is " + hawbLength + " characters (exceeds by " + (hawbLength - MAX_HAWB_LENGTH) + ")")
		hasInvalidHAWB = true
	} else {
		WebUI.comment(" Row ${i + 1} HAWB length is valid — [" + row.hawb + "] is " + hawbLength + " characters")
	}
}

if (!hasInvalidHAWB) {
	WebUI.comment(" WRONG TEST DATA — No HAWB exceeds max length of " + MAX_HAWB_LENGTH + ". Use a file with at least one HAWB longer than " + MAX_HAWB_LENGTH + " characters")
	assert false, "Wrong test file — no HAWB exceeds maximum length of " + MAX_HAWB_LENGTH
}

// Login
WebUI.openBrowser('')
WebUI.maximizeWindow()
WebUI.navigateToUrl('https://sf.ekonek.com/login')

WebUI.setText(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_Username'), GlobalVariable.Username)

WebUI.setEncryptedText(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_Password'), GlobalVariable.Password)

WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_Login'))



//  Upload Excel file
WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_Upload XLS'))

Robot robot = new Robot()
robot.delay(1000)
robot.keyPress(KeyEvent.VK_ESCAPE)
robot.keyRelease(KeyEvent.VK_ESCAPE)

WebUI.uploadFile(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_uploadxls'), filePath)
WebUI.comment(" File uploaded — " + excelRows.size() + " rows")



// Check system UI location
def dynamicXPath = { String xpath ->
	TestObject to = new TestObject()
	to.addProperty('xpath', ConditionType.EQUALS, xpath)
	return to
}

boolean systemAccepted = WebUI.waitForElementVisible(
	dynamicXPath("//div[@id='row-0']//div[@data-column-id='1']"), 5, FailureHandling.OPTIONAL)

if (systemAccepted) {
	
	WebUI.comment(" ROWS WITH INVALID HAWB ")

	// Show ALL rows that exceeded max length
	excelRows.eachWithIndex { row, i ->
		if (row.hawb.length() > MAX_HAWB_LENGTH) {
			String actualHAWB = WebUI.getText(dynamicXPath("//div[@id='row-${i}']//div[@data-column-id='1']"))
			WebUI.comment(" Row ${i + 1} — HAWB: [" + actualHAWB + "] | Length: " + actualHAWB.length() + " | Exceeds by: " + (actualHAWB.length() - MAX_HAWB_LENGTH) + " character(s)")
		}
	}
	WebUI.comment(" System accepts the file even HAWB exceeds its  maximum length of " + MAX_HAWB_LENGTH + " upon uploading")

	assert false, "TEST FAILED: System accepted file containing HAWB(s) exceeding max length of " + MAX_HAWB_LENGTH

} else {
	WebUI.comment(" System correctly REJECTED the file")
	WebUI.comment(" Data was NOT displayed in the table")
	WebUI.comment(" TEST PASSED — System rejects file with HAWB exceeding max length of " + MAX_HAWB_LENGTH)
}
