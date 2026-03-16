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

import java.awt.Robot
import java.awt.event.KeyEvent

String filePath = 'C:\\Users\\denlo\\OneDrive\\Desktop\\OJT Tasks\\OJT related tasks\\OJT MANUAL TESTING\\SF CHI\\SF-EXP-SUWP-00001.xlsx'


WebUI.openBrowser('')
WebUI.maximizeWindow()

WebUI.navigateToUrl('https://sf.ekonek.com/login')

WebUI.setText(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_Username'), 'NMM_User')

WebUI.setEncryptedText(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_Password'), 'IMrpfjBbSL8n+osp8It7RQ==')

WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_Login'))

WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_Upload XLS'))

// Close the native File Explorer dialog if it remains open after file selection
Robot robot = new Robot()
robot.delay(1000)
robot.keyPress(KeyEvent.VK_ESCAPE)
robot.keyRelease(KeyEvent.VK_ESCAPE)

WebUI.uploadFile(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_uploadxls'), filePath)


WebUI.comment("Excel file uploaded")

WebUI.delay(3)

WebUI.waitForElementVisible(findTestObject('Page_e-Konek Apps - SF Status Uploader/div_HAWB'), 10)


// Verify each header with a clear failure message
try {
	WebUI.verifyElementText(findTestObject('Page_e-Konek Apps - SF Status Uploader/div_HAWB'), 'HAWB')
	WebUI.comment(" HAWB column header verified")
} catch (Exception e) {
	WebUI.comment(" HAWB header mismatch — Wrong file may have been uploaded. Check the file: " + filePath)
	throw e
}

try {
	WebUI.verifyElementText(findTestObject('Page_e-Konek Apps - SF Status Uploader/div_EVENT CODE'), 'EVENT CODE')
	WebUI.comment(" EVENT CODE column header verified")
} catch (Exception e) {
	WebUI.comment(" EVENT CODE header mismatch — Wrong file may have been uploaded. Check the file: " + filePath)
	throw e
}

try {
	WebUI.verifyElementText(findTestObject('Page_e-Konek Apps - SF Status Uploader/div_REASON CODE'), 'REASON CODE')
	WebUI.comment(" REASON CODE column header verified")
} catch (Exception e) {
	WebUI.comment(" REASON CODE header mismatch — Wrong file may have been uploaded. Check the file: " + filePath)
	throw e
}

try {
	WebUI.verifyElementText(findTestObject('Page_e-Konek Apps - SF Status Uploader/div_DATE'), 'DATE')
	WebUI.comment(" DATE column header verified")
} catch (Exception e) {
	WebUI.comment(" DATE header mismatch — Wrong file may have been uploaded. Check the file: " + filePath)
	throw e
}

WebUI.comment("TEST PASSED — All UI table headers verified successfully")
