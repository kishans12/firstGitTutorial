package KeatPro;

import java.awt.Robot;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.io.IOException;
import java.net.URL;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.winium.DesktopOptions;
import org.openqa.selenium.winium.WiniumDriver;
import org.testng.TestListenerAdapter;
import org.testng.TestNG;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

public class Login 
	{	
		public static String[] getordno;
		public static double roundoff,cnvrtcaputil,placecnvrtcaputil,matchround;
		public static ExtentReports report1;
		public static ExtentTest log;
		public static WiniumDriver w;
		public static String caputilscrip,getscrip,caputil;
		WebElement keatpro,op;		
		
		public static void main(String[] args) 
		  {
			  TestListenerAdapter tla = new TestListenerAdapter();
			    TestNG testng = new TestNG();
			     Class[] classes = new Class[]
			     {
			    	Login.class		             
			     };
			     testng.setTestClasses(classes);
			     testng.addListener(tla);
			     testng.run();		  	  
		  	}
		
		String scrip1 = "AXISBANK~EQ";//"AXISBANK~EQ";
		Robot robot;

		public void captureutil() throws Exception
		{
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_R);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_P);//////////
			  Thread.sleep(500);
			  robot.keyRelease(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  
			  //count no of rows
			  WebElement nptable = w.findElement(By.name("Net Position")).findElement(By.id("33302"));			  
			  List < WebElement > rc = nptable.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			  System.out.println("row count "+ rc.size());
			  
			  for(int i=1;i<rc.size();i++)
			  {				  
				  WebElement abc = nptable.findElement(By.name("Row " + (i) + ""));
				  caputilscrip = abc.findElement(By.name("Row " + (i) + ", Column 1")).getText();
				  System.out.println("cap util "+caputilscrip);
				  if(scrip1.equals(caputilscrip))
				  {
					  caputil = abc.findElement(By.name("Row " + (i) + ", Column 18")).getText();
					  System.out.println("cap util "+caputil);
					  cnvrtcaputil = Double.parseDouble(caputil);
					  break;
				  }
				  else
				  {
					  caputil = "0.0";
					  System.out.println("cap util "+caputil);
					  System.out.println("Scrip not available to capture position");
				  }				  
			  }
			  
			  robot.keyPress(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_F);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_C);			  		  	
			  Thread.sleep(500);
			  robot.keyRelease(KeyEvent.VK_ALT);
			  Thread.sleep(500);
		}
		
		 @Test
		 public void test() throws IOException, Exception
		 {
			  DesktopOptions options= new DesktopOptions();
			  options.setApplicationPath("D:\\KeatProX_UAT\\KeatProVX.exe");
			  
			  w=new WiniumDriver(new URL("http://localhost:9999"),options);
			  
			  WebDriverWait wait= new WebDriverWait (w,2000);
			  WebElement element=wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("1006")));
			  element.click();			  
			  
			  //Thread.sleep(15000);
			  
			  //login to keatpro
			  w.findElement(By.id("1006")).sendKeys("1y039");
			  w.findElement(By.id("1007")).sendKeys("login1");
			  w.findElementById("1014").sendKeys("1111");
			  w.findElementByName("Login").click();
			  
			  WebDriverWait wait1= new WebDriverWait (w,3000);
			  WebElement element1=wait1.until(ExpectedConditions.visibilityOfElementLocated(By.name("Indices")));
			  
			  //Thread.sleep(35000);
			  
			  //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			  ///capture util
			  robot = new Robot();
			  			  
			 //String scrip1 = "AXISBANK~EQ";//"AXISBANK~EQ";
			  //table.find_elements_by_xpath("/*[@ControlType='Custom']/*[@ControlType='Custom'][2]")
			  	  
			 // WebElement wltable = w.findElement(By.id("65280")).findElement(By.id("33302"));
			 // WebElement colcount = wltable.findElement(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			 // System.out.println("colc ount " +colcount);
			 // List < WebElement > rc1 = wltable.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			 // System.out.println("row count "+ rc1.size());
			  			  
			  	
		    	//robot.keyPress(KeyEvent.VK_F5);
				//Thread.sleep(500);
				  
//////////////////////////////////////////////////////////////////////////////////
			  
	  		  captureutil();	
			  
/////////////////////////////////////////////////////////////			  
			  ///open watchlist
			  robot.keyPress(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_F);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_O);			  		  	
			  Thread.sleep(500);
			  robot.keyRelease(KeyEvent.VK_ALT);			  
			  
			  //extent report
			  report1 = new ExtentReports("D:\\testfile\\winapputil.html");
			  log = report1.startTest("Place/Modify/Cancel NSE Order");  
			  				
			  ///select watchlist
			  String WL = "UTILISATION";
			  w.findElement(By.id("1012")).click();
			  w.findElement(By.name("UTILISATION")).click();	  
			  w.findElement(By.id("1")).click();
			  log.log(LogStatus.INFO, "Watchlist opened");
			  
			  //keatpro = w.findElement(By.name("Watchlist - " + (WL) + "")).findElement(By.id("1000"));
			  
			  //count no of rows
			  WebElement wltable = w.findElement(By.id("65280")).findElement(By.id("1000"));
			  WebElement colcount = wltable.findElement(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			  System.out.println("colc ount " +colcount);
			  List < WebElement > rc1 = wltable.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			  System.out.println("row count "+ rc1.size());
			  
			  ///select scrip
			  for(int i=1;i<=rc1.size();i++)
			  {				  				  
				  getscrip = wltable.findElement(By.name("Row " + (i) + ", Column 1")).getText();
				  //System.out.println("get scrip "+getscrip);
				  
				  if(scrip1.equals(getscrip))
				  {
					  wltable.findElement(By.name("Row " + (i)+ ", Column 1")).click();
					  System.out.println("scrip got ");
					  log.log(LogStatus.INFO, "Scrip " +getscrip+ " is selected");

					  ///place order
					  break;
				  }
				  else
				  {
					  //log.log(LogStatus.ERROR, "Scrip not found");
					  System.out.println("scrip not found ");
				  }
			  }			  			
			  
			  ///place order
			  robot.keyPress(KeyEvent.VK_CONTROL);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_F1);
			  Thread.sleep(500);
			  
			  robot.keyRelease(KeyEvent.VK_CONTROL);
			  robot.keyRelease(KeyEvent.VK_F1);
			  Thread.sleep(500);
			  
			  robot.keyPress(KeyEvent.VK_DELETE);
			  Thread.sleep(500);
			  w.findElement(By.id("1350")).sendKeys("1");
			  robot.keyPress(KeyEvent.VK_TAB);
			  Thread.sleep(500);
			  String getamt = w.findElement(By.id("1409")).getText();
			  System.out.println("getamount "+getamt);
			  Double getamt1 = Double.parseDouble(getamt);
			  Double calc1 = getamt1 - 20.00;
			  Double calc = Math.round(calc1*100.0)/100.0;
			  if(calc<0)
			  {
				  calc=0.0;
			  }				
			  robot.keyPress(KeyEvent.VK_DELETE);
			  Thread.sleep(500);
			  w.findElement(By.id("1409")).sendKeys(String.valueOf(calc));
			  w.findElement(By.id("1360")).click();/////new buy order button
			  Thread.sleep(1000);			  			  
			  w.findElement(By.id("6")).click();//////confirmation yes
			  Thread.sleep(1000);
			  w.findElement(By.id("1")).click();//////success ok
			  Thread.sleep(800);
			  w.findElement(By.id("2")).click();//////success ok
			  Thread.sleep(500);
			  
			  ////get order no
			  String scrip2 = "Success. ABWS Order number is:";
			  for(int i=0;i<10;)
			  {					  				  
				  WebElement msg = w.findElementByClassName("SysListView32");
				  String getmsg = msg.findElement(By.xpath("//*[contains(@Name, '"+scrip2+"')]")).getAttribute("Name");///working
				  System.out.println("get message "+getmsg);
				  if(getmsg.startsWith(scrip2))
				  {
					  getordno = getmsg.split("\\:",0);
					  System.out.println("get on "+getordno[1]);
					  break;
				  }
				  else
				  {
					  System.out.println("No success message ");
				  }
				  
			  }		
			  log.log(LogStatus.INFO, "Order no " +getordno[1]+ " placed successfully");
			  System.out.println("Completed successfully");				  
			  	
			  ///cal reqd margin
			  Double calcutil = (calc * 1)/7.00;
			  roundoff = (Math.round(calcutil*100.0)/100.0);
			  System.out.println("calc util "+roundoff);				  
			  log.log(LogStatus.INFO, "Required Margin is : "+roundoff);
			  
			  //close utilisation
			  robot.keyPress(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_F);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_C);
			  Thread.sleep(500);
			  robot.keyRelease(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  
			  ///capture util
			  robot.keyPress(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_R);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_P);
			  Thread.sleep(500);
			  robot.keyRelease(KeyEvent.VK_ALT);
			  Thread.sleep(500);			  
			  
			  //count no of rows
			  WebElement nptable1 = w.findElement(By.name("Net Position")).findElement(By.id("33302"));			  
			  List < WebElement > rc2 = nptable1.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			  System.out.println("row count "+ rc2.size());
			  
			  for(int i=1;i<rc2.size();i++)
			  {
				  String placecaputilscrip = nptable1.findElement(By.name("Row " + (i) + ", Column 1")).getText();
				  System.out.println("place cap util scrip "+placecaputilscrip);
				  if(scrip1.equals(placecaputilscrip))
				  {
					  String placecaputil = nptable1.findElement(By.name("Row " + (i) + ", Column 18")).getText();
					  System.out.println("cap new util "+placecaputil);
					  placecnvrtcaputil = Double.parseDouble(placecaputil);
					  System.out.println("cnvrt cap util again "+cnvrtcaputil);
					  Double match = placecnvrtcaputil - cnvrtcaputil;
					  System.out.println("minus " +match);
					  matchround = Math.round(match*100.0)/100.0;
					  System.out.println("round minus " +matchround);						
					  log.log(LogStatus.INFO, "Margin utilised is : " +matchround);
					  
					  if(roundoff == matchround)
					  {
						  log.log(LogStatus.INFO, "Utilised");
						  System.out.println("utilisation is equal ");
					  }
					  else
					  {
						  log.log(LogStatus.INFO, "Not Utilised");
						  System.out.println("utilisation is unequal ");
					  }
					  
					  break;
				  }
				  else
				  {					  
					  System.out.println("Scrip not available to capture position");
				  }				  
			  }
			  
			  //close limits
			  robot.keyPress(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_F);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_C);
			  Thread.sleep(500);
			  robot.keyRelease(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  
			  //open order report
			  robot.keyPress(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_R);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_O);
			  Thread.sleep(500);
			  robot.keyRelease(KeyEvent.VK_ALT);
			  Thread.sleep(500);			  
			  
			  //keatpro =  w.findElement(By.name("Order Report [All]"));
		      		        
		      //count no of rows
			  WebElement ortable = w.findElement(By.name("Order Report [All]")).findElement(By.id("1003"));			  
			  List < WebElement > rc3 = ortable.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			  System.out.println("row count "+ rc3.size());
			  
			  ///calculate util for place			  
			  for(int i=1;i<=rc3.size();i++)
			  {				  				  
				  String capordno = ortable.findElement(By.name("Row " + (i) + ", Column 4")).getText();
				  System.out.println("get scrip "+capordno);				  
				  
					  ///modify order	
					  
					  String getstat = ortable.findElement(By.name("Row " + (i)+ ", Column 11")).getText();
					  System.out.println("status "+getstat);
					  
					  if(capordno.equals(getordno[1]) && getstat.equals("OPN"))
					  {		
						  ortable.findElement(By.name("Row " + (i)+ "")).click();
						  log.log(LogStatus.INFO, "Select order no : "+capordno);
						  Thread.sleep(500);
						  robot.mousePress(InputEvent.BUTTON3_DOWN_MASK);
						  Thread.sleep(500);
						  robot.mouseRelease(InputEvent.BUTTON3_DOWN_MASK);
						  Thread.sleep(500);
						  robot.keyPress(KeyEvent.VK_M);
						  Thread.sleep(500);
						  robot.keyPress(KeyEvent.VK_DELETE);
						  Thread.sleep(500);
						  w.findElement(By.id("1350")).sendKeys("3");
						  w.findElement(By.id("1360")).click();
						  Thread.sleep(1000);			  			  
						  w.findElement(By.id("6")).click();
						  Thread.sleep(1000);
						  w.findElement(By.id("1")).click();
						  Thread.sleep(1000);
						  w.findElement(By.id("2")).click();	///cancel modify dialog box		
						  log.log(LogStatus.INFO, "Order no " +getordno[1]+ " modified successfully");
						  
						  ///calc modified qty
						  Double RM = ((3 * calc ) / 7.00);
						  System.out.println("RM " +RM);
						  Double roundoff1 = (Math.round(RM*100.0)/100.0);
						  System.out.println("Modified Reqd margin is " +roundoff1);
						  log.log(LogStatus.INFO, "Modified Required Margin is : "+roundoff1);

						  ////chk modify util
						  //captureutil();
						  //System.out.println("matchround is "+matchround+ " cnvrtcaputil is "+cnvrtcaputil);
						  //Double utilreqmar = (Math.round((matchround - cnvrtcaputil)*100.0)/100.0);
						  //System.out.println("utilised req mar "+utilreqmar);
						  //log.log(LogStatus.INFO, "Utilised Margin is : "+utilreqmar);
						  						  
						  ///cancel order
						  ortable.findElement(By.name("Row " + (i)+ "")).click();
						  log.log(LogStatus.INFO, "Select order no : "+capordno);
						  Thread.sleep(500);
						  robot.mousePress(InputEvent.BUTTON3_DOWN_MASK);
						  Thread.sleep(500);
						  robot.mouseRelease(InputEvent.BUTTON3_DOWN_MASK);
						  Thread.sleep(500);
						  robot.keyPress(KeyEvent.VK_C);
						  Thread.sleep(500);
						  w.findElement(By.id("1360")).click();
						  Thread.sleep(1000);			  			  
						  w.findElement(By.id("6")).click();
						  Thread.sleep(1000);
						  w.findElement(By.id("1")).click();
						  Thread.sleep(1000);
						  w.findElement(By.id("2")).click();	///cancel cancel dialog box
						  log.log(LogStatus.INFO, "Order no " +getordno[1]+ " cancelled successfully");	
						  
						  captureutil();
						  System.out.println("cancel util "+caputil);
						  log.log(LogStatus.INFO, "Cancelled Required Margin is : "+caputil);

						  break;
					  }
					  else
					  {
						  //log.log(LogStatus.INFO, "Order is " +getstat+ " so cannot modify and cancel.");
						  System.out.println("Order is " +getstat+ " so cannot modify and cancel. ");
					  }	
			  }
			  			  
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
		 
		 
		 @Test(priority=4001)
			public void nouse()
			  {
				  report1.endTest(log);
				  report1.flush();
				  //w.get("D:\\Windowapp\\winapputil.html");				  
				  //w.get(srcfile);
			  }
		 
		}

