package KeatDealer;

import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.winium.DesktopOptions;
import org.openqa.selenium.winium.WiniumDriver;
import org.testng.TestListenerAdapter;
import org.testng.TestNG;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.Test;

//import com.library.variables;
import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import Main.beforesuitedealer;

public class TC1Randomchk extends beforesuitedealer
{	
	Robot robot;
	WebElement keatpro,op;		
	public static String scrip,getscrip,clientcode,action,getprice,enterscrip;
	int m,qty;
	public static double total,getamt1,roundoff;
	public static HSSFSheet s;
	public Row rw;
	public Cell cell;
	
	public static void main(String[] args) 
	  {
		  TestListenerAdapter tla = new TestListenerAdapter();
		    TestNG testng = new TestNG();
		     Class[] classes = new Class[]
		     {
		    	 TC1Randomchk.class
		     };
		     testng.setTestClasses(classes);
		     testng.addListener(tla);
		     testng.run();		  	  
	  	}
	
	//com.library.variables vr = new com.library.variables();	
	
	public void openWL() throws Exception
	{
			///open watchlist
			robot.keyPress(KeyEvent.VK_ALT);
			Thread.sleep(500);
			robot.keyPress(KeyEvent.VK_F);
			Thread.sleep(500);
			robot.keyPress(KeyEvent.VK_O);			  		  	
			Thread.sleep(500);
			robot.keyRelease(KeyEvent.VK_ALT);
			  
			///select watchlist
			String WL = s.getRow(1).getCell(6).getStringCellValue();
			w.findElement(By.id("1013")).click();
			WebElement wltable = w.findElement(By.name("WatchList")).findElement(By.id("1013"));
			wltable.findElement(By.xpath("//*[contains(@Name, '"+WL+"')]")).click();
			w.findElement(By.id("1")).click();
			log.log(LogStatus.INFO, "Watchlist opened");
			
			//count no of rows
			WebElement wltable1 = w.findElement(By.id("65281")).findElement(By.id("1000"));
			WebElement colcount = wltable1.findElement(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));		
			List < WebElement > rc1 = wltable1.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			System.out.println("row count "+ rc1.size());
			
			for(int k=1;k<=rc1.size();k++)
			  {
				  getscrip = wltable1.findElement(By.name("Row " + (k) + ", Column 1")).getText();
				  System.out.println("get scrip " +getscrip);
				  if(scrip.equals(getscrip))
				  {
					  wltable1.findElement(By.name("Row " + (k)+ ", Column 1")).click();
					  System.out.println("scrip got ");
					  getprice = wltable1.findElement(By.name("Row " + (k)+ ", Column 9")).getText();
					  System.out.println("get price "+getprice);
					  
					  //close watchlist
					  close();			  
					  break;
				  }
				  else
				  {
					  System.out.println("No scrip ");
				  }				  
			  }		  
	}
	
	public void buyorder() throws Exception
	{
		FileInputStream f = new FileInputStream(srcfile);
		wb = new HSSFWorkbook(f);		
		s = wb.getSheet("Randomchk");	
		  
		  for(int i=m;i<=m+5;i++)
		  {
			  Row row = s.getRow(i);
			  
			  String flag = s.getRow(i).getCell(0).getStringCellValue();
			  
			  if(flag.equals("Y"))
			  {
				  for(int j=1;j<row.getLastCellNum();)
				  {					  
							  robot.keyPress(KeyEvent.VK_TAB);
							  robot.keyPress(KeyEvent.VK_F1);
							  
							  ///select action
							  j=j+1;
							  action = s.getRow(i).getCell(j).getStringCellValue();
							  
							  if(action.equals("SELL"))							  
							  {
								  robot.keyPress(KeyEvent.VK_CONTROL);
								  robot.keyPress(KeyEvent.VK_F2);
								  robot.keyRelease(KeyEvent.VK_CONTROL);
								  robot.keyRelease(KeyEvent.VK_F2);
							  }
							  
							  ///enter client code
							  j=j+1;
							  clientcode = s.getRow(i).getCell(j).getStringCellValue();
							  //System.out.println("client code "+clientcode);
							  w.findElement(By.id("1018")).sendKeys(clientcode);							  
							  
							  robot.keyPress(KeyEvent.VK_TAB);
							  robot.keyPress(KeyEvent.VK_DELETE);
							  							  
							  ///enter security
							  j=j+1;
							  enterscrip = s.getRow(i).getCell(j).getStringCellValue();
							  w.findElement(By.id("1401")).sendKeys(enterscrip);
							  log.log(LogStatus.INFO, "Scrip " +enterscrip+ " is selected");
							  log.log(LogStatus.INFO, "Selected action is " +action);
							  robot.keyPress(KeyEvent.VK_TAB);
							  robot.keyPress(KeyEvent.VK_TAB);
							  
							  ///select nse
							  robot.keyPress(KeyEvent.VK_N);
							  
							  robot.keyPress(KeyEvent.VK_TAB);
							  robot.keyPress(KeyEvent.VK_TAB);
							  
							  ///enter qty
							  j=j+1;
							  qty = (int) s.getRow(i).getCell(j).getNumericCellValue();
							  w.findElement(By.id("1400")).sendKeys(String.valueOf(qty));
							  
							  robot.keyPress(KeyEvent.VK_TAB);
							  
							  if(!scrip.startsWith(enterscrip))
							  {
								  ///enter amt
								  System.out.println("scrip eaual "+scrip);
								  System.out.println("enterscrip eaual "+enterscrip);
								  getprice = "0.0";
							  }							  
							  
							  ///get price perc
							  j=j+1;
							  float perc = (float) s.getRow(i).getCell(j).getNumericCellValue();
							  getamt1 = (Double.parseDouble(getprice)*perc)/100;
							  double chngestrnggetamt = Double.parseDouble(getprice);
							  
							  if(action.equals("BUY"))
								{
									total = chngestrnggetamt-getamt1;
								}
								else
								{
									total = chngestrnggetamt+getamt1;
								}
							  roundoff = (double) (Math.round(total));
							  log.log(LogStatus.INFO, "Actual price is " +getprice);
							  log.log(LogStatus.INFO, "Modified price is " +roundoff);	
							  			
							  robot.keyPress(KeyEvent.VK_DELETE);
							  
							  if(roundoff<0)
							  {
								  roundoff=0.0;
							  }							  
							  w.findElement(By.id("1404")).sendKeys(String.valueOf(roundoff));
							  
							  j=j+1;
							  rw = s.getRow(i);
							  cell = rw.createCell(j);
							  cell.setCellValue(getprice);
							  
							  j=j+1;
							  rw = s.getRow(i);
							  cell = rw.createCell(j);
							  cell.setCellValue(roundoff);
							  
							  w.findElement(By.id("1342")).click();/////new buy order button
							  Thread.sleep(1500);
							  
							  //confirmation enter
							  robot.keyPress(KeyEvent.VK_ENTER);
							  Thread.sleep(1500);
							  //success enter
							  robot.keyPress(KeyEvent.VK_ENTER);
							  Thread.sleep(1000);
							  
							  robot.keyPress(KeyEvent.VK_TAB);
							  Thread.sleep(500);	
							  break;						  
				  }
			}
			else
		  	{
		  		System.out.println("Flag is N");
		  	}
		  }
		  
		  FileOutputStream Of = new FileOutputStream(srcfile);
			
			wb.write(Of);
			Of.close();		
	}
	
	public void close() throws Exception
	{
		  robot.keyPress(KeyEvent.VK_ALT);
		  Thread.sleep(500);
		  robot.keyPress(KeyEvent.VK_F);
		  Thread.sleep(500);
		  robot.keyPress(KeyEvent.VK_C);
		  Thread.sleep(500);
		  robot.keyRelease(KeyEvent.VK_ALT);
		  Thread.sleep(500);
	}
	
	public void ordrptf5() throws Exception
	{		  
		  robot.keyPress(KeyEvent.VK_ALT);
		  Thread.sleep(500);
		  robot.keyPress(KeyEvent.VK_R);
		  Thread.sleep(500);
		  robot.keyPress(KeyEvent.VK_O);
		  Thread.sleep(1000);	
		  robot.keyRelease(KeyEvent.VK_ALT);
		  Thread.sleep(1000);
		  robot.keyPress(KeyEvent.VK_TAB);
		  log.log(LogStatus.INFO, "Order Report refreshed ");
	}
	
	public void limitpgf5() throws Exception
	{		  		  
		  robot.keyPress(KeyEvent.VK_ALT);
		  Thread.sleep(500);
		  robot.keyPress(KeyEvent.VK_R);
		  Thread.sleep(500);
		  robot.keyPress(KeyEvent.VK_P);
		  Thread.sleep(500);
		  robot.keyRelease(KeyEvent.VK_ALT);
		  Thread.sleep(500);
		  log.log(LogStatus.INFO, "Open position refreshed ");
		  WebElement limittable = w.findElement(By.name("Net Position"));
		  limittable.findElement(By.id("1018")).sendKeys(clientcode);
		  robot.keyPress(KeyEvent.VK_F5);
		  Thread.sleep(500);
		  robot.keyPress(KeyEvent.VK_SHIFT);
		  Thread.sleep(500);
		  robot.keyPress(KeyEvent.VK_L);
		  Thread.sleep(500);
		  log.log(LogStatus.INFO, "Limits Report refreshed ");
		  robot.keyPress(KeyEvent.VK_D);
		  Thread.sleep(500);
		  log.log(LogStatus.INFO, "Today's position refreshed ");
		  robot.keyRelease(KeyEvent.VK_SHIFT);
	}
	
	public void trdrptf5() throws Exception
	{
		  robot.keyPress(KeyEvent.VK_ALT);
		  Thread.sleep(500);
		  robot.keyPress(KeyEvent.VK_R);
		  Thread.sleep(500);
		  robot.keyPress(KeyEvent.VK_T);
		  Thread.sleep(500);
		  robot.keyRelease(KeyEvent.VK_ALT);
		  Thread.sleep(500);
		  robot.keyPress(KeyEvent.VK_F5);
		  Thread.sleep(500);
		  log.log(LogStatus.INFO, "Trade Report refreshed ");
	}	
	
	@Test
	public void tcrandomchk() throws Exception, IOException
	{
			FileInputStream f = new FileInputStream(srcfile);
			wb = new HSSFWorkbook(f);
			
			s = wb.getSheet("Randomchk");
			
			///report creation
			report1 = new ExtentReports("D:\\testfile\\randomchk.html");
			log = report1.startTest("Randomly place order and check reports");
		  	
			////call login function
		  	Login();
		  	
		  	robot = new Robot();
		  	
		  	////call order report function
			ordrptf5();
			
			m=5;
			scrip = s.getRow(m).getCell(1).getStringCellValue();
			System.out.println("scrip "+scrip);
			openWL();
			buyorder();
			close();
			//ordrptf5();
			robot.keyPress(KeyEvent.VK_F5);
			trdrptf5();
			close();
			//limitpgf5();
			//close();
			
			m=m+6;
			System.out.println("m "+m);
			scrip = s.getRow(m).getCell(1).getStringCellValue();
			System.out.println("scrip "+scrip);
			openWL();
			buyorder();
			close();
			//ordrptf5();
			robot.keyPress(KeyEvent.VK_F5);
			trdrptf5();
			close();
			//limitpgf5();
			//close();
			
			m=m+6;
			System.out.println("m "+m);
			scrip = s.getRow(m).getCell(1).getStringCellValue();
			System.out.println("scrip "+scrip);
			openWL();
			close();
			ordrptf5();
			buyorder();
			close();
			robot.keyPress(KeyEvent.VK_F5);
			trdrptf5();
			close();
			limitpgf5();
			close();
			m=m-5;
			log.log(LogStatus.INFO, "Total orders placed "+m);
			
			
			///close project
			Thread.sleep(500);
			robot.keyPress(KeyEvent.VK_ALT);
			Thread.sleep(500);
			robot.keyPress(KeyEvent.VK_F);
			Thread.sleep(500);
			robot.keyPress(KeyEvent.VK_X);
			Thread.sleep(500);
			robot.keyRelease(KeyEvent.VK_ALT);
			Thread.sleep(500);
			w.findElement(By.id("6")).click();			  
			log.log(LogStatus.INFO, "Application closed");			
	}
	
	@Test(priority=201)
	public void nouse()
	  {
		  report1.endTest(log);
		  report1.flush();
		  //w.get("D:\\testfile\\randomchk.html");
		  //w.get(srcfile);
	  }
	
}
