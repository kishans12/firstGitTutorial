package com.TestCases;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

public class TC3Stockupload
{
	  public static WebDriver w1;
	  public static HSSFWorkbook wb;
	  public static HSSFSheet s2;
	  public static String srcfile = "D:\\Saleema\\KotakSecurities\\testdata.xls";
	  public static ExtentReports report1;
	  public static ExtentTest log;
	  String URL;
	  String clientacc;
	  String symbol;
	  int i,j,rc;
	  public static String stkupld;

	  com.library.Sellfromexistingutil sfeu = new com.library.Sellfromexistingutil();
	  
		@BeforeSuite
		public void beforeSuite() throws IOException
		{	  	
			  	System.setProperty("webdriver.chrome.driver", "D:\\Selenium software\\chromedriver.exe");
				w1 = new ChromeDriver();
				
				FileInputStream f = new FileInputStream(srcfile);
				wb = new HSSFWorkbook(f);
				s2 = wb.getSheet("Stockupload");
				
				/// get URL
				URL = s2.getRow(1).getCell(1).getStringCellValue();
				w1.get(URL);
				w1.manage().window().maximize();
				w1.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
				
		}
			
		@Test
		public void stckupload() throws Exception
		{
			FileInputStream f = new FileInputStream(srcfile);
			wb = new HSSFWorkbook(f);
			//FileOutputStream Of = new FileOutputStream(srcfile);
								
			s2 = wb.getSheet("Stockupload");
							
			///report creation
			report1 = new ExtentReports("D:\\Saleema\\stckupload.html");
			log = report1.startTest("Stock Upload");
			
			rc = s2.getLastRowNum() - s2.getFirstRowNum();
			System.out.println("rc "+rc);
				  	
			for(i=5;i<rc+1;i++)
			{
				Row row = s2.getRow(i);
		  		String flag = s2.getRow(i).getCell(0).getStringCellValue();
				
		  		if(flag.equals("Y"))
				{
					for(j=1;j<row.getLastCellNum();)
					{
						clientacc = s2.getRow(i).getCell(j).getStringCellValue();
						w1.findElement(By.id("clt_acc")).sendKeys(clientacc);
						log.log(LogStatus.INFO, "Entered client account is : " +clientacc);
						
						j=j+1;
						symbol = s2.getRow(i).getCell(j).getStringCellValue();
						w1.findElement(By.id("sec_symb")).sendKeys(symbol);
						log.log(LogStatus.INFO, "Entered symbol is : " +symbol);
						
						w1.findElement(By.xpath("//input[@value='Get Common Code']")).click();
						
						j=j+1;
						int stckbal = (int) s2.getRow(i).getCell(j).getNumericCellValue();
						w1.findElement(By.id("stk_bal")).sendKeys(String.valueOf(stckbal));
						log.log(LogStatus.INFO, "Entered stock balance is : " +stckbal);
						
						j=j+2;
						stkupld=s2.getRow(i).getCell(j).getStringCellValue();
	  					if(stkupld.equals("DEL"))
	  					{
	  						w1.findElement(By.xpath("//input[@value='DEL']")).click();
							log.log(LogStatus.INFO, "Click on delivery ");
	  					}
	  					else if(stkupld.equals("BNSTG"))
	  					{
	  						w1.findElement(By.xpath("//input[@value='BNSTG']")).click();
							log.log(LogStatus.INFO, "Click on delivery ");
	  					}
	  					else
	  					{
	  						w1.findElement(By.xpath("//input[@value='HOLD']")).click();
							log.log(LogStatus.INFO, "Click on delivery ");
	  					}
						
						String capturemsg = w1.findElement(By.xpath("//*[@id='container']/div/table/tbody/tr[2]/td/table/tbody/tr[1]/td")).getText();
						String actualmsg = "Successfully Uploaded!!";
						
						if(capturemsg.equals(actualmsg))
						{
							w1.findElement(By.partialLinkText("BACK")).click();
							log.log(LogStatus.INFO, "Stock " +symbol+ " got uploaded successfully");
							break;
						}
						else
						{
							log.log(LogStatus.INFO, "Unable to upload Stock as " +capturemsg );
							w1.close();
						}		
					}
				}
			}
					
			sfeu.beforeSuite1();
			sfeu.PlaceSFEOrder();
		}				

		@Test(priority=4003)
		public void nouse()
		  {
			  report1.endTest(log);
			  report1.flush();
			  w1.get("D:\\Saleema\\stckupload.html");
			  w1.get(srcfile);
		  }
 
}