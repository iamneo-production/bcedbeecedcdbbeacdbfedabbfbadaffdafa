package runner;

import java.io.FileInputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.time.Duration;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;

import org.openqa.selenium.interactions.Actions;
import org.apache.poi.ss.usermodel.DateUtil;
import org.openqa.selenium.support.events.WebDriverListener;
import org.testng.Assert;
import java.util.Set;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;
import utils.LoggerHandler;
import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import com.aventstack.extentreports.MediaEntityBuilder;
import org.openqa.selenium.JavascriptExecutor;
import java.time.Duration;
import java.util.concurrent.TimeUnit;
import java.util.ArrayList;
import java.util.List;
import javax.xml.crypto.Data;
import java.util.logging.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.Wait;

import utils.LoggerHandler;
import utils.Screenshot;
import utils.base64;

import utils.EventHandler;
import utils.Reporter;
public class Testcase1 extends Base {
    EventHandler e;
    java.util.logging.Logger log =  LoggerHandler.getLogger();
    base64 screenshotHandler = new base64();
    ExtentReports reporter = Reporter.generateExtentReport();;
     

@DataProvider(name = "testData")
    public Object[][] readTestData() throws IOException {
        String excelFilePath = System.getProperty("user.dir") + "/src/test/java/resources/Testdata.xlsx";
        String sheetName = "Sheet1";
    
        try (FileInputStream fileInputStream = new FileInputStream(excelFilePath);
             Workbook workbook = WorkbookFactory.create(fileInputStream)) {
    
            Sheet sheet = workbook.getSheet(sheetName);
            int rowCount = sheet.getLastRowNum();
    
            Object[][] searchDataArray = new Object[rowCount][5]; 
    
            for (int i = 1; i <= rowCount; i++) {
                Row row = sheet.getRow(i);
    
                searchDataArray[i - 1][0] = getStringCellValue(row.getCell(0));
                searchDataArray[i - 1][1] = getStringCellValue(row.getCell(1));
                searchDataArray[i - 1][2] = getStringCellValue(row.getCell(2));
                searchDataArray[i - 1][3] = getStringCellValue(row.getCell(3));
                searchDataArray[i - 1][4] = getStringCellValue(row.getCell(4));
               
            }
    
            return searchDataArray;
        }
    }
    
    private String getStringCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    double numericValue = cell.getNumericCellValue();
                    if (numericValue == (long) numericValue) {
                        return String.format("%d", (long) numericValue);
                    } else {
                        return String.valueOf(numericValue);
                    }
                }
            default:
                return "";
        }
    }
    

    
    @Test(priority = 1, dataProvider = "testData")
    public void addLowCostGiftCard(String recipientName, String recipientEmail, String senderName, String senderEmail, String Book ) throws InterruptedException, IOException {
        try {
            ExtentTest test = reporter.createTest("Add Low-Cost Gift Card", "Execution for adding a low-cost gift card");
            driver.get(prop.getProperty("url") + "/");
            log.info("Browser launched");
            driver.manage().window().maximize();
            
            driver.findElement(By.partialLinkText("Gift")).click();
            Select sortBy = new Select(driver.findElement(By.id("products-orderby")));
            sortBy.selectByVisibleText("Price: High to Low");
            log.info("List sorted");
            List<WebElement> searchResult = driver.findElements(By.xpath("//input[@value='Add to cart']"));
            searchResult.get(searchResult.size() - 1).click();
            log.info("******************");
            driver.findElement(By.id("giftcard_1_RecipientName")).sendKeys(recipientName);
            driver.findElement(By.id("giftcard_1_RecipientEmail")).sendKeys(recipientEmail);
            driver.findElement(By.id("giftcard_1_SenderName")).sendKeys(senderName);
            driver.findElement(By.id("giftcard_1_SenderEmail")).sendKeys(senderEmail);
            driver.findElement(By.id("add-to-cart-button-1")).click();
            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10)); 
            WebElement cartQtySpan = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[@class='cart-qty']")));
            Assert.assertTrue(cartQtySpan.getText().contains("1"));
            test.pass("Test passed successfully");

            
        } catch (Exception ex) {
            ex.printStackTrace(); 
            ExtentTest test = reporter.createTest("addLow Cost GiftCard");
            Screenshot.getScreenShot("add LowCost GiftCard");
            String base64Screenshot = screenshotHandler.captureScreenshotAsBase64(driver);
            test.log(Status.FAIL, "Unable to Perform the add LowCost GiftCard", MediaEntityBuilder.createScreenCaptureFromBase64String(base64Screenshot).build());
            }
    }


    
    @Test(priority = 2, dataProvider = "testData")
    public void TestCheckButton(String recipientName, String recipientEmail, String senderName, String senderEmail, String Book ) throws InterruptedException, IOException {
        try {
            ExtentTest test = reporter.createTest("TestCheckButton", "Execution for Checkout Button");
            driver.get(prop.getProperty("url") + "/");
            log.info("Browser launched");
            driver.manage().window().maximize();
            log.info("******************");
            driver.findElement(By.partialLinkText("Gift")).click();
            Select sortBy = new Select(driver.findElement(By.id("products-orderby")));
            sortBy.selectByVisibleText("Price: High to Low");
            
            List<WebElement> searchResult = driver.findElements(By.xpath("//input[@value='Add to cart']"));
            searchResult.get(searchResult.size() - 1).click();
            
            driver.findElement(By.id("giftcard_1_RecipientName")).sendKeys(recipientName);
            driver.findElement(By.id("giftcard_1_RecipientEmail")).sendKeys(recipientEmail);
            driver.findElement(By.id("giftcard_1_SenderName")).sendKeys(senderName);
            driver.findElement(By.id("giftcard_1_SenderEmail")).sendKeys(senderEmail);
            driver.findElement(By.id("add-to-cart-button-1")).click();
            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10)); 
            driver.findElement(By.xpath("//span[contains(text(),'Shopping cart')]")).click();
            Assert.assertFalse(driver.findElement(By.id("termsofservice")).isSelected());
            driver.findElement(By.xpath("//input[@id='termsofservice']")).click();
            driver.findElement(By.xpath("//button[@id='checkout']")).click();
            test.pass("Test passed successfully");

            
        } catch (Exception ex) {
            ex.printStackTrace(); 
            ExtentTest test = reporter.createTest("Test Checkout Button");
            Screenshot.getScreenShot("Test Checkout Butto");
            String base64Screenshot = screenshotHandler.captureScreenshotAsBase64(driver);
            }
    }
     @Test(priority = 3, dataProvider = "testData")
    
     public void TestCheckout(String recipientName, String recipientEmail, String senderName, String senderEmail, String Book ) throws InterruptedException, IOException {
            try {
                ExtentTest test = reporter.createTest("TestCheckoutButton", "Execution for Checkoutout Button");
                driver.get(prop.getProperty("url") + "/");
                log.info("Browser launched");
                driver.manage().window().maximize();
                
                driver.findElement(By.partialLinkText("Gift")).click();
                Select sortBy = new Select(driver.findElement(By.id("products-orderby")));
                sortBy.selectByVisibleText("Price: High to Low");
                
                List<WebElement> searchResult = driver.findElements(By.xpath("//input[@value='Add to cart']"));
                searchResult.get(searchResult.size() - 1).click();
                
                driver.findElement(By.id("giftcard_1_RecipientName")).sendKeys(recipientName);
                driver.findElement(By.id("giftcard_1_RecipientEmail")).sendKeys(recipientEmail);
                driver.findElement(By.id("giftcard_1_SenderName")).sendKeys(senderName);
                driver.findElement(By.id("giftcard_1_SenderEmail")).sendKeys(senderEmail);
                driver.findElement(By.id("add-to-cart-button-1")).click();
                WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10)); 
                driver.findElement(By.xpath("//span[contains(text(),'Shopping cart')]")).click();
                Assert.assertFalse(driver.findElement(By.id("termsofservice")).isSelected());
                driver.findElement(By.xpath("//input[@id='termsofservice']")).click();
                driver.findElement(By.xpath("//button[@id='checkout']")).click();
                driver.findElement(By.xpath("//body/div[4]/div[1]/div[4]/div[2]/div[1]/div[2]/div[1]/div[1]/div[3]/input[1]")).click();
                driver.findElement(By.id("BillingNewAddress_FirstName")).sendKeys("Kevin");
                driver.findElement(By.id("BillingNewAddress_LastName")).sendKeys("De bruyne");
                driver.findElement(By.id("BillingNewAddress_Email")).sendKeys("Debruyne@macity.com");
                Select country = new Select(driver.findElement(By.xpath("//select[@id='BillingNewAddress_CountryId']")));
                country.selectByVisibleText("United States");
                driver.findElement(By.id("BillingNewAddress_City")).sendKeys("Ny city");
                log.info("******************");
                driver.findElement(By.id("BillingNewAddress_Address1")).sendKeys("NY");
                driver.findElement(By.id("BillingNewAddress_ZipPostalCode")).sendKeys("23201");
                driver.findElement(By.id("BillingNewAddress_PhoneNumber")).sendKeys("8983039");
                driver.findElement(By.xpath("/html[1]/body[1]/div[4]/div[1]/div[4]/div[1]/div[1]/div[2]/ol[1]/li[1]/div[2]/div[1]/input[1]")).click();
                driver.findElement(By.xpath("/html[1]/body[1]/div[4]/div[1]/div[4]/div[1]/div[1]/div[2]/ol[1]/li[2]/div[2]/div[1]/input[1]")).click();
                driver.findElement(By.xpath("/html[1]/body[1]/div[4]/div[1]/div[4]/div[1]/div[1]/div[2]/ol[1]/li[3]/div[2]/div[1]/input[1]")).click();
                driver.findElement(By.xpath("/html[1]/body[1]/div[4]/div[1]/div[4]/div[1]/div[1]/div[2]/ol[1]/li[4]/div[2]/div[2]/input[1]")).click();
                driver.findElement(By.xpath("//a[contains(text(),'Click here for order details.')]")).click();
                test.pass("Test passed successfully");  

            } catch (Exception ex) {
                ExtentTest test = reporter.createTest("Test Checkout exception");
                Screenshot.getScreenShot("TestCheckout");
                String base64Screenshot = screenshotHandler.captureScreenshotAsBase64(driver);
                test.log(Status.FAIL, "Unable to Perform the add Checkout", MediaEntityBuilder.createScreenCaptureFromBase64String(base64Screenshot).build());
               
            }
        }
    @Test(priority = 4, dataProvider = "testData")
    
    public void Booksearch(String recipientName, String recipientEmail, String senderName, String senderEmail, String Book ) throws InterruptedException, IOException {
        try {
                   ExtentTest test = reporter.createTest("TestCheckButton", "Execution for Checkout Button");
                   driver.get(prop.getProperty("url") + "/");
                   log.info("Browser launched");
                   driver.manage().window().maximize();
                   driver.findElement(By.id("small-searchterms")).sendKeys(Book);
                   driver.findElement(By.xpath("//input[@value='Search']")).click();
                   WebElement searchInput = driver.findElement(By.id("Q"));
                   String inputValue = searchInput.getAttribute("value");
                   Assert.assertTrue(inputValue.contains("book"), "Search input value contains 'book'");
                   test.pass("Test passed successfully");  
   
               } catch (Exception ex) {
                ex.printStackTrace(); 
                ExtentTest test = reporter.createTest("Book search Exception");
                Screenshot.getScreenShot("Book search");
                String base64Screenshot = screenshotHandler.captureScreenshotAsBase64(driver);
                test.log(Status.FAIL, "Unable to Perform Book search", MediaEntityBuilder.createScreenCaptureFromBase64String(base64Screenshot).build());
             }
           }

    
@BeforeMethod
public void beforeMethod() throws MalformedURLException {
    openBrowser();

}

    @AfterMethod
    public void afterMethod() {
        driver.quit();
        reporter.flush();
        log.info("Browser closed");
        LoggerHandler.closeHandler();
    }
}