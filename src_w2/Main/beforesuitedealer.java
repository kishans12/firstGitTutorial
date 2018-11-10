package Main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
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

import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Image;
import com.itextpdf.text.pdf.PdfWriter;
import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import Libraries.variables;

public class beforesuitedealer 
{
	public static ExtentReports report1;
	public static ExtentTest log;
	public static WiniumDriver w;
	public static String srcfile = "D:\\testfile\\winapptestdata.xls";
	public static HSSFWorkbook wb;
	public static HSSFSheet s1;
	public static String URL, Uname, Pwd;
	  public static int a;
	  public com.itextpdf.text.Document doc;
	  public static int abc=0;
		public String forreportscripname;

	
	@BeforeSuite
	  public void beforeSuite() throws IOException
	  {
		  FileInputStream f = new FileInputStream(srcfile);
		  wb = new HSSFWorkbook(f);
		  s1 = wb.getSheet("Config");
		  
		  URL = s1.getRow(3).getCell(1).getStringCellValue();
		  DesktopOptions options= new DesktopOptions();
		  options.setApplicationPath(URL);
		  
		  w=new WiniumDriver(new URL("http://localhost:9999"),options);
		  
////////////////////////////////////////////////////////////////////////////	
		  //comment this 2 line for UAT///
		  //WebElement dialogbox = w.findElement(By.className("#32770"));
		  //dialogbox.findElement(By.id("6")).click();
////////////////////////////////////////////////////////////////////////////		  
		  
		  WebDriverWait wait= new WebDriverWait (w,2000);
		  WebElement element=wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("1006")));
		  element.click();		 		  
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
		  File failedtest = new File("D:\\testfile\\failedscrn.jpg");
		  FileUtils.copyFile(snapshot, failedtest);
	  }   
	  
	  public void deltakesnapshot() throws IOException
	  {
		  File index = new File("D:\\testfile\\passscrns");
		  
		  String[]entries = index.list();
		  for(String s: entries)
		  {
		      File currentFile = new File(index.getPath(),s);
		      currentFile.delete();
		  }	 
		  a=0;
	  }
	  
	  public void passtakesnapshot() throws IOException
	  {	  
		  File snapshot = ((TakesScreenshot)w).getScreenshotAs(OutputType.FILE);
		  a++;	  
		  File passtest = new File("D:\\testfile\\passscrns\\" +a+ ".jpg");
		  FileUtils.copyFile(snapshot, passtest);
	  }
	  
	  public void copy2pdf() throws IOException, DocumentException
	  {	  
		  doc = new com.itextpdf.text.Document();	  
		  PdfWriter.getInstance(doc, new FileOutputStream("D:\\testfile\\passscrns\\mock.pdf"));
		  doc.open();
		  System.out.println("a "+a);
		  for(int i=1;i<=a;i++)
		  {
			  Image img = Image.getInstance("D:\\testfile\\passscrns\\" +i+ ".jpg");	  
			  //img.setAbsolutePosition(500f, 650f);
			  //img.scalePercent(40, 40);
			  img.scaleToFit(500, 500);
			  doc.add(img);
		  }	   	    	  
		  doc.close();  
		  
		  //////copy pdf file	  
		  abc++;
		  File source = new File("D:\\testfile\\passscrns\\mock.pdf");
		  File dest = new File("D:\\testfile\\copypdffile\\Mock.pdf");
		  FileUtils.copyFile(source, dest);
	  }
	  
	  @AfterMethod
	  public void aftermethod(ITestResult res) throws Exception
	  {
		  if(res.getStatus()==ITestResult.FAILURE)
		  {
			  log.log(LogStatus.FAIL,"Failed");
			  takesnapshot();
			  res.getThrowable();
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
