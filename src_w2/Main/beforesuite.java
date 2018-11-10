package Main;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.URL;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.winium.DesktopOptions;
import org.openqa.selenium.winium.WiniumDriver;
import org.testng.ITestResult;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeSuite;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

public class beforesuite 
{
	public static ExtentReports report1;
	public static ExtentTest log;
	public static WiniumDriver w;
	public static String srcfile = "D:\\testfile\\winapptestdata.xls";
	public static HSSFWorkbook wb;
	public static HSSFSheet s1;
	public static String Uname, Pwd;
	
	@BeforeSuite
	  public void beforeSuite() throws IOException
	  {
		  DesktopOptions options= new DesktopOptions();
		  options.setApplicationPath("D:\\KSDealerX_UAT\\KSDealerVX.exe");
		  
		  w=new WiniumDriver(new URL("http://localhost:9999"),options);
		  
		  WebDriverWait wait= new WebDriverWait (w,2000);
		  WebElement element=wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("1006")));
		  element.click();
		  
		  FileInputStream f = new FileInputStream(srcfile);
		  wb = new HSSFWorkbook(f);
		  s1 = wb.getSheet("Config");
	  }
	
	  public void Login() throws IOException
	  {
		  //login to dealer
		  
		  ///enter username
		  Uname = s1.getRow(3).getCell(2).getStringCellValue();
		  w.findElement(By.id("1006")).sendKeys(Uname);
		  log.log(LogStatus.INFO, "Enter User name " +Uname);
		  	
		  ///enter password
		  Pwd  = s1.getRow(3).getCell(3).getStringCellValue();
		  w.findElement(By.id("1007")).sendKeys(Pwd);
		  log.log(LogStatus.INFO, "Enter Password ******");
			
		  w.findElementByName("Login").click();
		  log.log(LogStatus.INFO, "Logged in successfully..");
		  
		  WebDriverWait wait1= new WebDriverWait (w,3000);
		  WebElement element1=wait1.until(ExpectedConditions.visibilityOfElementLocated(By.name("Indices")));
	  }
	  
	  public void takesnapshot() throws IOException
	  {
		  File snapshot = ((TakesScreenshot)w).getScreenshotAs(OutputType.FILE);
		  File failedtest = new File("D:\\Winapp\\failedscreenshot.jpg");
		  FileUtils.copyFile(snapshot, failedtest);
	  }
	  
	  @AfterMethod
	  public void aftermethod(ITestResult res) throws Exception
	  {
		  if(res.getStatus()==ITestResult.FAILURE)
		  {
			  log.log(LogStatus.FAIL,"Failed");
			  takesnapshot();
		  }
		  else if(res.getStatus()==ITestResult.SKIP)
		  {
			log.log(LogStatus.SKIP,"Skipped");
		  }
		  else if(res.getStatus()==ITestResult.SUCCESS)
		  {
			log.log(LogStatus.PASS,"Passed");
		  }
		  
	  }
	  
}
