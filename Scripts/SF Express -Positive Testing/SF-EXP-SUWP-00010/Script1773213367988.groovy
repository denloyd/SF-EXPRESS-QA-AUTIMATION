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

String filePath = 'C:\\Users\\denlo\\OneDrive\\Desktop\\OJT Tasks\\OJT related tasks\\OJT MANUAL TESTING\\SF CHI\\SF-EXP-SUWP-00010.xlsx'


// Read Excel data BEFORE uploading
FileInputStream fis = new FileInputStream(filePath)
XSSFWorkbook workbook = new XSSFWorkbook(fis)
def sheet = workbook.getSheetAt(0)
DataFormatter formatter = new DataFormatter()

// Store all rows
List<Map<String, String>> excelRows = []
for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
	def row = sheet.getRow(rowNum)
	if (row == null) continue

	String hawb       = formatter.formatCellValue(row.getCell(0)).trim()
	String eventCode  = formatter.formatCellValue(row.getCell(1)).trim()
	String reasonCode = formatter.formatCellValue(row.getCell(2)).trim()
	String date       = formatter.formatCellValue(row.getCell(3)).trim()

	// Skip completely empty rows
	if (hawb.isEmpty() && eventCode.isEmpty() && reasonCode.isEmpty() && date.isEmpty()) continue

	excelRows.add([
		hawb      : hawb,
		eventCode : eventCode,
		reasonCode: reasonCode,
		date      : date
	])
}

workbook.close()
fis.close()

WebUI.comment("Total rows to be saved: " + excelRows.size())
excelRows.eachWithIndex { row, i ->
	WebUI.comment("Row ${i + 1} — HAWB: ${row.hawb} | EVENT: ${row.eventCode} | REASON: ${row.reasonCode} | DATE: ${row.date}")
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
WebUI.comment(" Excel file uploaded — " + excelRows.size() + " rows")
WebUI.delay(3)

// EVENT_CODE + REASON_CODE validation
WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_Save'))

WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_Proceed'))


// accepted valid EVENT_CODE + REASON_CODE
WebUI.waitForElementVisible(findTestObject('Page_e-Konek Apps - SF Status Uploader/div_Saved successfully'), 15)
WebUI.verifyElementText(findTestObject('Page_e-Konek Apps - SF Status Uploader/div_Saved successfully'), 'Saved successfully.')
WebUI.comment(" TEST PASSED — System accepted valid EVENT_CODE + REASON_CODE combination")

