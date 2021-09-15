package ExpressAutomation;


import com.relevantcodes.extentreports.LogStatus;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import org.openqa.selenium.By;

import org.openqa.selenium.WebElement;

import org.testng.annotations.Test;

import java.io.File;
import java.io.IOException;
import java.util.List;

public class ExpressWebsite extends Reusable_Annotations_Class_Html_Report {

    @Test
    public void ExpressWebsite () throws InterruptedException, BiffException, IOException, WriteException {
        Workbook readableFile = Workbook.getWorkbook(new File("src/main/resources/expressSheet.xls"));
        WritableWorkbook writableFile = Workbook.createWorkbook(new File("src/main/resources/expressSheet_results.xls"),readableFile);
        WritableSheet writableSheet = writableFile.getSheet(0);


        int rows = writableSheet.getRows();

        for(int i = 1; i < rows; i++) {
            driver.navigate().to("https://www.express.com");
            Thread.sleep(2500);
            if(i == 1) {
                WebElement closeBtn = ReusableMethods_With_Logger.getWebElement(driver,"//*[@id='closeModal']",logger);
                ReusableMethods_With_Logger.click(driver, closeBtn,"Close popup", logger);
            }
            WebElement mensClothingHover = ReusableMethods_With_Logger.getWebElement(driver,"//*[@href='/mens-clothing']",logger);
            ReusableMethods_With_Logger.mouseHover(driver,mensClothingHover,"Mouse hover mens clothing section",logger);
            Thread.sleep(1000);
            WebElement mensPoloShirtsSection = ReusableMethods_With_Logger.getWebElement(driver,"//*[@href='/mens-clothing/shirts/polos/cat1006']",logger);
            ReusableMethods_With_Logger.click(driver, mensPoloShirtsSection,"click on mens polo section",logger);
            Thread.sleep(2000);
            WebElement firstClothingSelection = ReusableMethods_With_Logger.getIndexedWebElement(driver, "//*[@class='_2fbIe3xmE78JEQRb26pdpQ']", 0,logger);
            ReusableMethods_With_Logger.click(driver,firstClothingSelection,"click on first clothing option",logger);
            Thread.sleep(2000);
            WebElement clothingSize = ReusableMethods_With_Logger.getWebElement(driver,"//*[@class='_29GwyLL9tJIABAZ0llJMdA'] [@value='"+writableSheet.getCell(0,i).getContents() +"']",logger);
            ReusableMethods_With_Logger.click(driver, clothingSize,"select clothing size",logger);
            Thread.sleep(1000);
            WebElement addToBagBtn = ReusableMethods_With_Logger.getWebElement(driver,"//*[text()='Add to Bag']",logger);
            ReusableMethods_With_Logger.click(driver,addToBagBtn,"click add to bag",logger);
            Thread.sleep(2000);
            WebElement viewBagBtn = ReusableMethods_With_Logger.getWebElement(driver,"//*[text()='View Bag']",logger);
            ReusableMethods_With_Logger.click(driver,viewBagBtn,"click on view bag button",logger);
            Thread.sleep(3000);
            WebElement quantityDropDown = ReusableMethods_With_Logger.getWebElement(driver,"//*[@id='qdd-0-quantity']",logger);
            ReusableMethods_With_Logger.selectDropdownByValue(driver, quantityDropDown, writableSheet.getCell(1,i).getContents(),"Select quantity of item",logger);
            Thread.sleep(1000);
            WebElement continueToCheckoutBtn = ReusableMethods_With_Logger.getWebElement(driver,"//*[@id='continue-to-checkout']",logger);
            ReusableMethods_With_Logger.click(driver,continueToCheckoutBtn ,"click on continue to checkout",logger);
            Thread.sleep(2000);
            WebElement checkoutAsGuestBtn = ReusableMethods_With_Logger.getWebElement(driver,"//*[text()='Checkout as Guest']",logger);
            ReusableMethods_With_Logger.click(driver, checkoutAsGuestBtn,"click on checkout as guest button",logger);
            Thread.sleep(3000);
            WebElement firstNameTextBox = ReusableMethods_With_Logger.getWebElement(driver,"//*[@id='contact-information-firstname']",logger);
            WebElement lastNameTextBox = ReusableMethods_With_Logger.getWebElement(driver,"//*[@name='lastname']",logger);
            WebElement emailTextBox = ReusableMethods_With_Logger.getWebElement(driver,"//*[@name='email']",logger);
            WebElement confirmEmailTextBox = ReusableMethods_With_Logger.getWebElement(driver,"//*[@name='confirmEmail']",logger);
            WebElement phoneNumberTextBox = ReusableMethods_With_Logger.getWebElement(driver,"//*[@name='phone']",logger);

            ReusableMethods_With_Logger.sendKeys(driver, firstNameTextBox, writableSheet.getCell(2,i).getContents(),"enter first name in text box",logger);
            ReusableMethods_With_Logger.sendKeys(driver, lastNameTextBox, writableSheet.getCell(3,i).getContents(),"enter last name in text box",logger);
            ReusableMethods_With_Logger.sendKeys(driver, emailTextBox, writableSheet.getCell(4,i).getContents(),"enter email in text box",logger);
            ReusableMethods_With_Logger.sendKeys(driver, confirmEmailTextBox, writableSheet.getCell(4,i).getContents(),"enter confirm email in text box",logger);
            ReusableMethods_With_Logger.sendKeys(driver, phoneNumberTextBox, writableSheet.getCell(5,i).getContents(),"enter phone number in text box",logger);

            WebElement checkoutContinueBtn = ReusableMethods_With_Logger.getWebElement(driver,"//*[text()='Continue']",logger);
            ReusableMethods_With_Logger.click(driver,checkoutContinueBtn ,"click on continue" ,logger);
            Thread.sleep(3000);

           /*
            WebDriverWait wait = new WebDriverWait(driver,7);
            try{
                wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@name='bluecoreCloseButton']")));
                ReusableMethods_With_Logger.click(driver, "//*[@name='bluecoreCloseButton']", logger);
            }catch (TimeoutException e)
            {
                continue;
            }*/

            WebElement shippingLineTextBox = ReusableMethods_With_Logger.getWebElement(driver,"//*[@name='shipping.line1']",logger);
            WebElement postalCodeTextBox = ReusableMethods_With_Logger.getWebElement(driver,"//*[@name='shipping.postalCode']",logger);
            WebElement cityTextBox = ReusableMethods_With_Logger.getWebElement(driver,"//*[@name='shipping.city']",logger);
            WebElement stateDropdown = ReusableMethods_With_Logger.getWebElement(driver,"//*[@name='shipping.state']",logger);

            ReusableMethods_With_Logger.sendKeys(driver,shippingLineTextBox , writableSheet.getCell(6,i).getContents(),"enter address in text box" ,logger);
            ReusableMethods_With_Logger.sendKeys(driver, postalCodeTextBox, writableSheet.getCell(7,i).getContents(),"enter zip code in text box",logger);
            ReusableMethods_With_Logger.sendKeys(driver, cityTextBox, writableSheet.getCell(8,i).getContents(),"enter city in text box",logger);
            ReusableMethods_With_Logger.selectDropdownByValue(driver, stateDropdown, writableSheet.getCell(9,i).getContents(),"select state from dropdown",logger);

            WebElement continueBtn2 = ReusableMethods_With_Logger.getWebElement(driver,"//*[text()='Continue']",logger);

            ReusableMethods_With_Logger.click(driver, continueBtn2,"click on continue",logger);
            Thread.sleep(3000);
            WebElement continueBtn3 = ReusableMethods_With_Logger.getWebElement(driver,"//*[text()='Continue']",logger);
            ReusableMethods_With_Logger.click(driver,continueBtn3 , "click on continue",logger);
            Thread.sleep(1000);


            WebElement elt = ReusableMethods_With_Logger.getWebElement(driver, "//*[@class='_1Q4iDku_IopeC7OgnKsdoD']",logger);
            Label label = null;
            try {
                List<WebElement> childrenElts = elt.findElements(By.xpath("./child::*"));
                label = new Label(10, i, childrenElts.get(1).getText() + " " + childrenElts.get(2).getText());
                logger.log(LogStatus.PASS,"able to retrieve text");
            }catch (Exception e)
            {
                logger.log(LogStatus.FAIL,"unable to retrieve text: " + e);
            }
            writableSheet.addCell(label);
            driver.manage().deleteAllCookies();
        }
        Thread.sleep(2000);

        writableFile.write();
        writableFile.close();

        driver.close();
    }
}
