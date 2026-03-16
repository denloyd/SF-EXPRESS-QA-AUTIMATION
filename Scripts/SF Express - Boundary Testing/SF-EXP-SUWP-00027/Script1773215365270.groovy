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

String filePath = 'C:\\Users\\denlo\\OneDrive\\Desktop\\OJT Tasks\\OJT related tasks\\OJT MANUAL TESTING\\SF CHI\\SF-EXP-SUWP-00027.xlsx'


// Read ALL EVENT_CODE values from Excel BEFORE uploading
FileInputStream fis = new FileInputStream(filePath)
XSSFWorkbook workbook = new XSSFWorkbook(fis)
def sheet = workbook.getSheetAt(0)
DataFormatter formatter = new DataFormatter()

// store EVENT_CODE
List<Map<String, String>> excelRows = []
for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
	def row = sheet.getRow(rowNum)
	if (row == null) continue

	def eventCodeCell = row.getCell(1)  
	if (eventCodeCell == null) continue
	String eventCodeValue = formatter.formatCellValue(eventCodeCell).trim()
	if (eventCodeValue.isEmpty()) continue

	excelRows.add([
		eventCode : eventCodeValue,
		rowNum    : String.valueOf(rowNum + 1)
	])
}

workbook.close()
fis.close()

WebUI.comment("Total rows read from Excel: " + excelRows.size())


// Check ALL EVENT_CODE values against max length
int MAX_EVENTCODE_LENGTH = 7  

boolean hasInvalidEventCode = false

excelRows.eachWithIndex { row, i ->
	int eventCodeLength = row.eventCode.length()
	WebUI.comment("Row ${i + 1} — EVENT_CODE: [" + row.eventCode + "] | Length: " + eventCodeLength + " | Max: " + MAX_EVENTCODE_LENGTH)

	if (eventCodeLength > MAX_EVENTCODE_LENGTH) {
		WebUI.comment(" Row ${i + 1} EVENT_CODE EXCEEDS max length — [" + row.eventCode + "] is " + eventCodeLength + " characters (exceeds by " + (eventCodeLength - MAX_EVENTCODE_LENGTH) + ")")
		hasInvalidEventCode = true
	} else {
		WebUI.comment(" Row ${i + 1} EVENT_CODE length is valid — [" + row.eventCode + "] is " + eventCodeLength + " characters")
	}
}

if (!hasInvalidEventCode) {
	WebUI.comment(" WRONG TEST DATA — No EVENT_CODE exceeds max length of " + MAX_EVENTCODE_LENGTH + ". Use a file with at least one EVENT_CODE longer than " + MAX_EVENTCODE_LENGTH + " characters")
	assert false, "Wrong test file — no EVENT_CODE exceeds maximum length of " + MAX_EVENTCODE_LENGTH
}



// Login
WebUI.openBrowser('')
WebUI.maximizeWindow()
WebUI.navigateToUrl('https://sf.ekonek.com/login')

WebUI.setText(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_Username'), 'NMM_User')
WebUI.setEncryptedText(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_Password'), 'IMrpfjBbSL8n+osp8It7RQ==')
WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_Login'))



// Upload Excel file
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
	
	WebUI.comment(" ROWS WITH INVALID EVENT_CODE ")

	// Show ALL rows that exceeded max length
	excelRows.eachWithIndex { row, i ->
		if (row.eventCode.length() > MAX_EVENTCODE_LENGTH) {
			
			String actualEventCode = WebUI.getText(dynamicXPath("//div[@id='row-${i}']//div[@data-column-id='2']"))
			WebUI.comment(" Row ${i + 1} — EVENT_CODE: [" + actualEventCode + "] | Length: " + actualEventCode.length() + " | Exceeds by: " + (actualEventCode.length() - MAX_EVENTCODE_LENGTH) + " character(s)")
		}
	}

	WebUI.comment("System accepts the uploaded file even the EVENT_CODE exceeds its maximum length of " + MAX_EVENTCODE_LENGTH + " upon uploading")

	assert false, "TEST FAILED: System accepted file containing EVENT_CODE exceeding max length of " + MAX_EVENTCODE_LENGTH

} else {
	WebUI.comment(" System correctly REJECTED the file")
	WebUI.comment(" Data was NOT displayed in the table")
	WebUI.comment(" TEST PASSED — System rejects file with EVENT_CODE exceeding max length of " + MAX_EVENTCODE_LENGTH)
}