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
import com.kms.katalon.core.testobject.ConditionType as ConditionType
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable
import org.openqa.selenium.Keys as Keys
import java.awt.Robot
import java.awt.event.KeyEvent

String filePath = 'C:\\Users\\denlo\\OneDrive\\Desktop\\OJT Tasks\\OJT related tasks\\OJT MANUAL TESTING\\SF CHI\\SF-EXP-SUWP-00028.xlsx'


WebUI.openBrowser('')
WebUI.maximizeWindow()
WebUI.navigateToUrl('https://sf.ekonek.com/login')

// Login
WebUI.setText(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_Username'), 'NMM_User')
WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_Password'))
WebUI.setEncryptedText(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_Password'), 'IMrpfjBbSL8n+osp8It7RQ==')
WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_Login'))
WebUI.delay(3)

// Upload file
WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_Upload XLS'))
Robot robot = new Robot()
robot.delay(1000)
robot.keyPress(KeyEvent.VK_ESCAPE)
robot.keyRelease(KeyEvent.VK_ESCAPE)
WebUI.uploadFile(findTestObject('Page_e-Konek Apps - SF Status Uploader/input_uploadxls'), filePath)
WebUI.delay(3)

// Click Save
WebUI.click(findTestObject('Page_e-Konek Apps - SF Status Uploader/button_Save'))
WebUI.delay(2)

// Read the actual error message from the popup
TestObject popupMessage = new TestObject('popupMessage')
popupMessage.addProperty('xpath', ConditionType.EQUALS, "//div[@id='swal2-html-container']//li")
WebUI.waitForElementVisible(popupMessage, 10)

String actualMessage   = WebUI.getText(popupMessage).trim()
String expectedMessage = 'Invalid Event Code'

WebUI.comment('Expected message: [' + expectedMessage + '] | Actual message: [' + actualMessage + ']')


// Verify the error message
WebUI.comment("Upon saving, the system rejects the uploaded file but it does not  display proper error message")
WebUI.verifyMatch(actualMessage, expectedMessage, false)

