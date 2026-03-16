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
import java.awt.Robot
import java.awt.event.KeyEvent

WebUI.openBrowser('')
WebUI.maximizeWindow()
WebUI.navigateToUrl('https://sf.ekonek.com/login')

// Helper — dynamic XPath
def dynamicXPath = { String xpath ->
	TestObject to = new TestObject()
	to.addProperty('xpath', ConditionType.EQUALS, xpath)
	return to
}



WebUI.comment('TEST 1: HAWB exceeds enhancement requirement limit of 20 characters.')

WebUI.setText(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_Username'), 'NMM_User')
WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_Password'))
WebUI.setEncryptedText(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_Password'), 'IMrpfjBbSL8n+osp8It7RQ==')
WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_Login'))

WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_Upload XLS'))

// Close the native File Explorer dialog if it remains open after file selection
Robot robot = new Robot()
robot.delay(1000)
robot.keyPress(KeyEvent.VK_ESCAPE)
robot.keyRelease(KeyEvent.VK_ESCAPE)

WebUI.uploadFile(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_uploadxls'),
	'C:\\Users\\denlo\\OneDrive\\Desktop\\OJT Tasks\\OJT related tasks\\OJT MANUAL TESTING\\SF CHI\\SF-EXP-SUWP-00026v2.xlsx')
WebUI.delay(3)

WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_Save'))
WebUI.delay(2)

// Check what happened after Save
boolean test1ErrorPopup  = WebUI.waitForElementVisible(dynamicXPath("//div[@id='swal2-html-container']//li"), 5, FailureHandling.OPTIONAL)
boolean test1SaveSuccess = WebUI.waitForElementVisible(dynamicXPath("//button[normalize-space()='Proceed']"),  5, FailureHandling.OPTIONAL)

WebUI.comment('TEST 1 RESULT ')
if (test1SaveSuccess) {
	WebUI.comment(' System saved the data to the database.')
	WebUI.comment(' No proper error message was displayed.')
	WebUI.comment(' System should have rejected and displayed a proper error message.')
	// Continue the flow since system saved
	WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_Proceed'))
	WebUI.delay(2)
	
} else if (test1ErrorPopup) {
	String test1Message = WebUI.getText(dynamicXPath("//div[@id='swal2-html-container']//li")).trim()
	WebUI.comment(' PASS: System rejected the file — error popup appeared.')
	WebUI.comment('Actual error message: [' + test1Message + ']')
	WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_OK'))
}

WebUI.comment('TEST 1 — HAWB exceeds requirement limit (20 chars):')
WebUI.comment(test1SaveSuccess
	? 'FAILED: Upon saving, the system accepts the HAWB even it exceeds to its maximum limit based on the Enhancement Requirement'
	: ' PASS: System rejected — verify error message is correct')

// TEST 1 ASSERTION — should NOT save to DB (will FAIL because of the bug)
WebUI.verifyEqual(test1SaveSuccess, false, FailureHandling.CONTINUE_ON_FAILURE)

WebUI.delay(5)


WebUI.comment('TEST 2: HAWB exceeds current database max length of 50 characters.')

WebUI.navigateToUrl('https://sf.ekonek.com/login')
WebUI.setText(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_Username'), 'NMM_User')
WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_Password'))
WebUI.setEncryptedText(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_Password'), 'IMrpfjBbSL8n+osp8It7RQ==')
WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_Login'))

WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_Upload XLS'))

// Close the native File Explorer dialog if it remains open after file selection

robot.delay(1000)
robot.keyPress(KeyEvent.VK_ESCAPE)
robot.keyRelease(KeyEvent.VK_ESCAPE)

WebUI.uploadFile(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_uploadxls'),
	'C:\\Users\\denlo\\OneDrive\\Desktop\\OJT Tasks\\OJT related tasks\\OJT MANUAL TESTING\\SF CHI\\SF-EXP-SUWP-00026.xlsx')
WebUI.delay(3)

WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_Save'))
WebUI.delay(2)

// Check what happened after Save
boolean test2ErrorPopup  = WebUI.waitForElementVisible(dynamicXPath("//div[@id='swal2-html-container']//li"), 5, FailureHandling.OPTIONAL)
boolean test2SaveSuccess = WebUI.waitForElementVisible(dynamicXPath("//button[normalize-space()='Proceed']"),  5, FailureHandling.OPTIONAL)

WebUI.comment('TEST 2 RESULT')
if (test2ErrorPopup) {
	String test2Message = WebUI.getText(dynamicXPath("//div[@id='swal2-html-container']//li")).trim()
	WebUI.comment(' PASS: System rejected the file — error popup appeared.')
	WebUI.comment('Actual error message: [' + test2Message + ']')

	if (test2Message.contains('ORA-12899')) {
		WebUI.comment(' BUG: System is exposing a backend Oracle DB error to the user.')
		WebUI.comment(' BUG: Actual   : [' + test2Message + ']')
		WebUI.comment(' BUG: Expected : A proper user-friendly message such as:')
		WebUI.comment(' BUG:            "HAWB exceeds the maximum allowed length of 50 characters."')
	} else {
		WebUI.comment(' MANUAL CHECK: Is this a proper user-friendly error message?')
	}
	WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_OK'))
} else if (test2SaveSuccess) {
	WebUI.comment(' BUG: System saved the data to the database instead of rejecting.')
	WebUI.comment(' BUG: No error message was displayed.')
	WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_Proceed'))
	WebUI.delay(2)
	
}


boolean test2ShowsOracleError = test2ErrorPopup &&
	WebUI.getText(dynamicXPath("//div[@id='swal2-html-container']//li")).trim().contains('ORA-12899')

WebUI.verifyEqual(test2SaveSuccess,      false, FailureHandling.CONTINUE_ON_FAILURE)  // should not save
WebUI.verifyEqual(test2ShowsOracleError, false, FailureHandling.CONTINUE_ON_FAILURE)  // should not show Oracle error


WebUI.comment('TEST 2 — HAWB exceeds DB limit (50 chars):')
WebUI.comment(test2ShowsOracleError
	? 'FAILED: Upon saving, system rejects the uploaded file but it displays incorrect error message'
	: test2SaveSuccess
		? 'FAILED: Upon saving, system rejects the uploaded file but it displays incorrect error message'
		: ' PASS: System rejected with proper message')

WebUI.delay(5)
WebUI.closeBrowser()