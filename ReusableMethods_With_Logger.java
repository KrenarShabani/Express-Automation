package ExpressAutomation;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;
import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.commons.io.FileUtils;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.File;

public class ReusableMethods_With_Logger {

    public static WebDriver setDriver()
    {
        WebDriverManager.chromedriver().setup();
        ChromeOptions options = new ChromeOptions();
        options.addArguments("start-maximized");
        options.addArguments("incognito");
        WebDriver webDriver = new ChromeDriver(options);
        return webDriver;

    }

    public static WebElement getWebElement(WebDriver driver, String xpath,ExtentTest logger)
    {
        WebElement webElt = null;
        try{webElt = driver.findElement(By.xpath(xpath));
            logger.log(LogStatus.PASS,"Able to retrieve webElement " + xpath);

        }
        catch (Exception e)
        {
            System.out.println("Couldn't find webElement: " + xpath + " " + e);
            logger.log(LogStatus.FAIL,"Unable to retrieve webElement " + xpath + " " + e);
        }
        return webElt;
    }

    public static WebElement getIndexedWebElement(WebDriver driver, String xpath, int index,ExtentTest logger)
    {
        WebElement elt = null;
        try
        {
           elt = driver.findElements(By.xpath(xpath)).get(index);
           logger.log(LogStatus.PASS,"Able to retrieve indexed webElement " + xpath);
        }catch (Exception e)
        {
            System.out.println("Couldn't find webElement array at: "+xpath+ " " + e);
            logger.log(LogStatus.FAIL,"Unable to retrieve indexed webElement " + xpath + " " + e);

        }
        return elt;
    }

    //select dropdown
    public static void selectDropdownByIndex (WebDriver driver, WebElement elt, String index,String Action, ExtentTest logger)
    {
        try {
            Select selector = new Select(elt);
            selector.selectByIndex(Integer.parseInt(index));
            logger.log(LogStatus.PASS,"Able to " + Action);
        } catch (Exception e)
        {
            System.out.println("couldn't " + Action + ": " + e);
            logger.log(LogStatus.FAIL,"unable to " + Action + ": " + e);
        }
    }
    public static void selectDropdownByValue (WebDriver driver, WebElement elt, String Value,String Action, ExtentTest logger)
    {
        try {
            Select selector = new Select(elt);
            selector.selectByValue(Value);
            logger.log(LogStatus.PASS,"Able to " + Action);
        } catch (Exception e)
        {
            System.out.println("couldn't " + Action + ": " + e);
            logger.log(LogStatus.FAIL,"unable to " + Action + ": " + e);
        }
    }

    public static String getText(WebDriver driver, WebElement elt,String Action, ExtentTest logger)
    {
        String text = "";
        try
        {
            text = elt.getText();
            logger.log(LogStatus.PASS,"Able to retrieve text from webElement");
        } catch(Exception e){
            System.out.println("couldn't " + Action + ": " + e);
            logger.log(LogStatus.FAIL,"unable to " + Action + ": " + e);
        }
        return text;
    }
    //get text

    public static void sendKeys(WebDriver driver,WebElement elt, String text,String Action,ExtentTest logger)
    {
        try
        {
            elt.sendKeys(text);
            logger.log(LogStatus.PASS,"Able to " + Action);
        }catch (Exception e)
        {
            System.out.println("couldn't " + Action + ": " + e);
            logger.log(LogStatus.FAIL,"unable to " + Action + ": " + e);
        }

    }

    public static void click(WebDriver driver, WebElement elt,String Action, ExtentTest logger)
    {
        try
        {
            elt.click();
            logger.log(LogStatus.PASS,"Able to " + Action);
        } catch (Exception e)
        {
            System.out.println("couldn't " + Action + ": " + e);
            logger.log(LogStatus.FAIL,"unable to " + Action + ": " + e);
        }

    }

    public static void mouseHover(WebDriver driver, WebElement elt,String Action,ExtentTest logger)
    {
        try {
            Actions action = new Actions(driver);
            action.moveToElement(elt).perform();
            logger.log(LogStatus.PASS,"Able to " + Action);
        }catch (Exception e)
        {
            System.out.println("couldn't " + Action + ": " + e);
            logger.log(LogStatus.FAIL,"unable to " + Action + ": " + e);

        }
    }

    public static void scroll(WebDriver driver, int from, int to,ExtentTest logger)
    {
        try {
            JavascriptExecutor jse = (JavascriptExecutor) driver;
            jse.executeScript("scroll(" + from + "," + to + ")");
            logger.log(LogStatus.PASS,"Scroll successful");
        }catch (Exception e)
        {
            System.out.println("couldn't scroll");
            logger.log(LogStatus.FAIL,"unable to scroll");
        }
    }
    //slider by SendKeys method
    public static void sliderSendKeysMethod(WebDriver driver, String xpath, int resetPoint, int desiredPoint, String elementName) throws InterruptedException {
        WebDriverWait wait = new WebDriverWait(driver, 15);

        try {

            //storing the WebElement as a variable
            WebElement slider = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(xpath)));

            //calling mouse action to click on the slider
            Actions action = new Actions(driver);
            action.click(slider).build().perform();
            Thread.sleep(1500);

            //reset loop
            for (int i = 0; i < resetPoint; i++) {
                action.sendKeys(Keys.ARROW_LEFT).build().perform();
            }//end of reset

            Thread.sleep(1500);

            //desired loop
            for (int i = 0; i < desiredPoint; i++) {
                action.sendKeys(Keys.ARROW_RIGHT).build().perform();
            }//end of desired
            Thread.sleep(1500);

        } catch (Exception e) {

            System.out.println("Unable to move the slider " +elementName + " " + e);

        }//end of slider by SendKeys method



        }
    public static void verifyPageTitle(WebDriver driver, String expectedTitle, ExtentTest logger){
        //verifying the title using if else condition
        String actualTitle = driver.getTitle();
        if(actualTitle.equals(expectedTitle)){
            System.out.println("Title matches");
            logger.log(LogStatus.PASS,"Title matches");
        } else {
            System.out.println("Title doesn't match " + actualTitle);
            logger.log(LogStatus.FAIL,"Title doesn't match " + actualTitle);
        }//end of get title condition
    }//end of verifyPageTitle
    public static void getScreenShot(WebDriver driver,String imageName, ExtentTest logger) {
        try {
            String fileName = imageName + ".png";
            String directory = null;
            String snPath = null;
            directory = "src/main/java/Extent_Report/Screenshots/";
            snPath = "Screenshots//";
            File sourceFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
            FileUtils.copyFile(sourceFile, new File(directory + fileName));
            //String imgPath = directory + fileName;
            String image = logger.addScreenCapture(snPath + fileName);
            logger.log(LogStatus.PASS, "ScreenShot Taken", image);
        } catch (Exception e) {
            logger.log(LogStatus.FAIL, "Error Occured while taking screenshot");
            e.printStackTrace();
        }
    }//end of getScreenshot method

}
