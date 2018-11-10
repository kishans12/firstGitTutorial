package Main;

import org.testng.annotations.BeforeSuite;
import org.testng.asserts.SoftAssert;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.testng.ITestResult;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;

public class BSOrderXLIE 
{
  public static WebDriver w;
  public static ExtentReports report1;
  public static ExtentTest log;
  public static HSSFSheet s2;
  String URL;
  String Uname;
  String Pwd;
  int AccCode;
  public static String srcfile = "D:\\Saleema\\KotakSecurities\\testdata.xls";
  public static HSSFWorkbook wb;
    
  @BeforeSuite
  public void beforeSuite() throws IOException
  {	  	
	  	System.setProperty("webdriver.ie.driver", "D:\\Selenium software\\drivers\\IEDriverServer.exe");
		w = new InternetExplorerDriver();
		
		FileInputStream f = new FileInputStream(srcfile);
		wb = new HSSFWorkbook(f);
		s2 = wb.getSheet("Config");
		
		/// get URL
		URL = s2.getRow(2).getCell(1).getStringCellValue();
		w.get(URL);
		//w.manage().window().maximize();
		w.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		
  }
  
  public static void checkurl()
  {
	  //to chk for login page error
	  /*
	  String imagetxt = "Login to your Trading Account"; 
	  String match = w.findElement(By.xpath("//*[@id=\'container2\']/div/form/h2")).getText();
	  SoftAssert asser = new SoftAssert();
	  asser.assertEquals(imagetxt, match);
	  //asser.assertAll();
	  //to check for logged in page error
	  String imagettx2 = "";
	  String match2 = w.findElement(By.xpath("//*[@id=\'KSEC-WelcomeText\']")).getText();
	  SoftAssert asser2 = new SoftAssert();
	  asser2.assertEquals(imagettx2, match2);
	  */
	  String imagetxt = "Welcome ";
	  String match = w.findElement(By.xpath("//*[@id=\'KSEC-WelcomeText\']")).getText();
	  System.out.println("match text "+match);
	  if(imagetxt.startsWith(match))
	  {
		  String capturemsg = w.findElement(By.xpath("//*[@id=\'container\']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
		  log.log(LogStatus.ERROR, "" +capturemsg);
	  }
	  
  }
  
  public void Login() throws IOException
  {	  		
		///enter username
	  	Uname = s2.getRow(2).getCell(2).getStringCellValue();
	  	w.findElement(By.id("userid")).sendKeys(Uname);
	  	log.log(LogStatus.INFO, "Enter User name " +Uname);
	  	
	  	///enter password
	  	Pwd  = s2.getRow(2).getCell(3).getStringCellValue();
		w.findElement(By.id("pwd")).sendKeys(Pwd);
		log.log(LogStatus.INFO, "Enter Password ******");
		
		///enter access code
		AccCode= (int) s2.getRow(2).getCell(4).getNumericCellValue();
		w.findElement(By.id("otp")).sendKeys(String.valueOf(AccCode));
		log.log(LogStatus.INFO, "Enter Security Key/Access Code ****");
		
		///click on login
		//w.findElement(By.partialLinkText("Login")).click();
		w.findElement(By.xpath("//*[@id='tbdiv']/span")).click();
///////////////////xtra added
		//w.findElement(By.xpath("//*[@id='tbdiv']/span")).click();
//////////////////
		System.out.println("Logged in Successfully");
		log.log(LogStatus.INFO, "Logged in Successfully");
		//*[@id="login"]/div[5]/div[5]/button/span
		//checkurl();
  }
   
  public void takesnapshot() throws IOException
  {
	  File snapshot = ((TakesScreenshot)w).getScreenshotAs(OutputType.FILE);
	  File failedtest = new File("D:\\Saleema\\KotakSecurities\\src\\NormalType.jpg");
	  FileUtils.copyFile(snapshot, failedtest);
  }
  
  @AfterMethod
  public void aftermethod(ITestResult res) throws Exception
  {
	  if(res.getStatus()==ITestResult.FAILURE)
	  {
		  log.log(LogStatus.FAIL,"Failed");
		  takesnapshot();
		  //System.out.println("TC Failed");
		  //res.getThrowable();
	  }
	  else if(res.getStatus()==ITestResult.SKIP)
	  {
		log.log(LogStatus.SKIP,"Skipped");
		//System.out.println("TC Skipped");
	  }
	  else if(res.getStatus()==ITestResult.SUCCESS)
	  {
		log.log(LogStatus.PASS,"Passed");
		//System.out.println("TC Success");
	  }
	  
  }
 
  @AfterSuite
  public void afterSuite() 
  {		
	  	
		//logout
		//w.findElement(By.partialLinkText("Logout")).click();
		//log.log(LogStatus.INFO, "Application logged out successfully");	
		
		//w.get("D:\\Saleema\\AMOCashreport.html");
		//close browser
		
		//w.close();
		//log.log(LogStatus.INFO, "Browser closed");
		//System.out.println("Closed");
		
  }

}
