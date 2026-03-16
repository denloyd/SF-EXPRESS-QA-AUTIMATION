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

String filePath = 'C:\\Users\\denlo\\OneDrive\\Desktop\\OJT Tasks\\OJT related tasks\\OJT MANUAL TESTING\\SF CHI\\SF-EXP-SUWP-00029.xlsx'


// Read ALL REASON_CODE values from Excel BEFORE uploading
FileInputStream fis = new FileInputStream(filePath)
XSSFWorkbook workbook = new XSSFWorkbook(fis)
def sheet = workbook.getSheetAt(0)
DataFormatter formatter = new DataFormatter()

//store REASON_CODE 
List<Map<String, String>> excelRows = []
for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
	def row = sheet.getRow(rowNum)
	if (row == null) continue

	def reasonCodeCell = row.getCell(2)  
	if (reasonCodeCell == null) continue
	String reasonCodeValue = formatter.formatCellValue(reasonCodeCell).trim()
	if (reasonCodeValue.isEmpty()) continue

	excelRows.add([
		reasonCode : reasonCodeValue,
		rowNum     : String.valueOf(rowNum + 1)
	])
}

workbook.close()
fis.close()


WebUI.comment("Total rows read from Excel: " + excelRows.size())


// Check ALL REASON_CODE values against max length
int MAX_REASONCODE_LENGTH = 7  

boolean hasInvalidReasonCode = false

excelRows.eachWithIndex { row, i ->
	int reasonCodeLength = row.reasonCode.length()
	WebUI.comment("Row ${i + 1} — REASON_CODE: [" + row.reasonCode + "] | Length: " + reasonCodeLength + " | Max: " + MAX_REASONCODE_LENGTH)

	if (reasonCodeLength > MAX_REASONCODE_LENGTH) {
		WebUI.comment(" Row ${i + 1} REASON_CODE EXCEEDS max length — [" + row.reasonCode + "] is " + reasonCodeLength + " characters (exceeds by " + (reasonCodeLength - MAX_REASONCODE_LENGTH) + ")")
		hasInvalidReasonCode = true
	} else {
		WebUI.comment(" Row ${i + 1} REASON_CODE length is valid — [" + row.reasonCode + "] is " + reasonCodeLength + " characters")
	}
}

if (!hasInvalidReasonCode) {
	WebUI.comment(" WRONG TEST DATA — No REASON_CODE exceeds max length of " + MAX_REASONCODE_LENGTH + ". Use a file with at least one REASON_CODE longer than " + MAX_REASONCODE_LENGTH + " characters")
	assert false, "Wrong test file — no REASON_CODE exceeds maximum length of " + MAX_REASONCODE_LENGTH
}


// Login
WebUI.openBrowser('')
WebUI.maximizeWindow()
WebUI.navigateToUrl('https://sf.ekonek.com/login')

WebUI.setText(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_Username'), 'NMM_User')
WebUI.setEncryptedText(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_Password'), 'IMrpfjBbSL8n+osp8It7RQ==')
WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_Login'))


//  Upload Excel file
WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_Upload XLS'))

Robot robot = new Robot()
robot.delay(1000)
robot.keyPress(KeyEvent.VK_ESCAPE)
robot.keyRelease(KeyEvent.VK_ESCAPE)

WebUI.uploadFile(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_uploadxls'), filePath)
WebUI.comment(" File uploaded — " + excelRows.size() + " rows")
WebUI.delay(3)


// Check system behavior
def dynamicXPath = { String xpath ->
	TestObject to = new TestObject()
	to.addProperty('xpath', ConditionType.EQUALS, xpath)
	return to
}

boolean systemAccepted = WebUI.waitForElementVisible(
	dynamicXPath("//div[@id='row-0']//div[@data-column-id='1']"), 5, FailureHandling.OPTIONAL)

if (systemAccepted) {
	WebUI.comment(" ROWS WITH INVALID REASON_CODE ")

	// Show ALL rows that exceeded max length
	excelRows.eachWithIndex { row, i ->
		if (row.reasonCode.length() > MAX_REASONCODE_LENGTH) {
			
			String actualReasonCode = WebUI.getText(dynamicXPath("//div[@id='row-${i}']//div[@data-column-id='3']"))
			WebUI.comment(" Row ${i + 1} — REASON_CODE: [" + actualReasonCode + "] | Length: " + actualReasonCode.length() + " | Exceeds by: " + (actualReasonCode.length() - MAX_REASONCODE_LENGTH) + " character(s)")
		}
	}

	WebUI.comment(" System accepts the uploaded file even the EVENT_CODE exceeds its maximum length of " + MAX_REASONCODE_LENGTH + " upon uploading")

	assert false, "TEST FAILED: System accepted file containing REASON_CODE exceeding max length of " + MAX_REASONCODE_LENGTH

} else {
	WebUI.comment(" System correctly REJECTED the file")
	WebUI.comment(" Data was NOT displayed in the table")
	WebUI.comment(" TEST PASSED — System rejects file with REASON_CODE exceeding max length of " + MAX_REASONCODE_LENGTH)
}