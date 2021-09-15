package ExpressAutomation;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import org.openqa.selenium.WebDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeSuite;

import java.io.IOException;
import java.lang.reflect.Method;

public class Reusable_Annotations_Class_Html_Report {

    public static WebDriver driver;
    public static ExtentReports reports;
    public static ExtentTest logger;

    @BeforeSuite
    public void setTheDriver() throws IOException{
        driver = ReusableMethods_With_Logger.setDriver();
        reports = new ExtentReports("src/main/java/Extent_report/AutomationReport/report.html",true);
    }
    @BeforeMethod
    public void getTestName(Method testname)
    {
        logger = reports.startTest(testname.getName());
    }

    @AfterMethod
    public void endtest()
    {
        reports.endTest(logger);
    }
    @AfterSuite
    public void quitSession(){
        reports.flush();
        driver.quit();
    }

}
