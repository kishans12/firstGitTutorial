/*
package KeatPro;

import java.awt.Robot;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.URL;

import org.apache.poi.hssf.record.RightMarginRecord;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.winium.DesktopOptions;
import org.openqa.selenium.winium.WiniumDriver;
import org.testng.annotations.Test;

public class vitloop 
{
	  public static String srcfile = "D:\\Saleema\\Windowproj\\testdata.xls";
	  public static HSSFWorkbook wb;
	  public static HSSFSheet s2;

	 @Test
	 public void test() throws IOException, Exception
	 {
		  DesktopOptions options= new DesktopOptions();
		  options.setApplicationPath("D:\\KeatProX_UAT\\KeatProVX.exe");
		  
		  WiniumDriver w=new WiniumDriver(new URL("http://localhost:9999"),options);
			//w.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		  
		  Thread.sleep(15000);
		  
			FileInputStream f = new FileInputStream(srcfile);
			wb = new HSSFWorkbook(f);
			s2 = wb.getSheet("Config");
			
			String Uname = s2.getRow(2).getCell(2).getStringCellValue();
		  	w.findElement(By.id("1006")).sendKeys(Uname);
		  	
		  	///enter password
		  	String Pwd  = s2.getRow(2).getCell(3).getStringCellValue();
			w.findElement(By.id("1007")).sendKeys(Pwd);
			
			///enter access code
			int AccCode= (int) s2.getRow(2).getCell(4).getNumericCellValue();
			w.findElement(By.id("1014")).sendKeys(String.valueOf(AccCode));
		  	
			w.findElementByName("Login").click();
			
			Thread.sleep(25000);			   	
			
			//w.findElement(By.id("59648")).click();
			///create watchlist			
			  Robot robot = new Robot();
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_F);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_N);			  		  	
			  Thread.sleep(500);
			  robot.keyRelease(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  robot.mousePress(InputEvent.BUTTON3_DOWN_MASK);
			  Thread.sleep(500);
			  robot.mouseRelease(InputEvent.BUTTON3_DOWN_MASK);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_W);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_R);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_TAB);
			  Thread.sleep(500);
			  robot.keyRelease(KeyEvent.VK_TAB);
			  Thread.sleep(500);
			  String watchlistname = "SUPMULTIPLE";//s2.getRow(2).getCell(2).getStringCellValue();
			  w.findElement(By.id("1015")).sendKeys(watchlistname);
			  w.findElement(By.id("1")).click();
			  
			  //check for input error
			  String getmsg = w.findElement(By.id("Input Error")).getText();
			  System.out.println("get msg "+getmsg);
			  
	 }
}
*/

/*
 				//count no of rows
			  WebElement mytable = w.findElement(By.id("65280")).findElement(By.id("33302"));			  
			  List < WebElement > rc = mytable.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			  System.out.println("row count "+ rc.size());
			  			  
			 			  
		    	for (int row1 = 1; row1 < rc.size(); row1++) 
		    	{
					  WebElement mytable1 = w.findElement(By.id("65280")).findElement(By.id("33302")).findElement(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
					  
		    	   // WebElement Columns_row = mytable.findElement(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
					List < WebElement > colcnt = mytable1.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
					
		    	    System.out.println("col count "+colcnt.size());
		    	    
		    	    System.out.println("Number of cells In Row " + row1 + " are " + colcnt.size());
		    	    
		    	    //Loop will execute till the last cell of that specific row.
		    	    for (int column = 0; column < colcnt.size(); column++) 
		    	    {
		    	        // To retrieve text from that specific cell.
						//getscrip = colcnt.get(row1).findElement(By.name("Row " + (row1+1) + ", Column 1")).getText();
			    	    //System.out.println("Cell Value of row number " + row1 + " and column number " + column + " Is " + getscrip);
		    	    	//getscrip = mytable.findElement(By.name("Row " + (row1) + ", Column 1")).getText(); //correct
		    	    	//System.out.println("Cell Value of row number " + row1 + " and column number " + column + " Is " + getscrip);
		    	    	
		    	        String scrip = colcnt.get(column).getText();
		    	        System.out.println(".............Cell Value of row number " + row1 + " and column number " + column + " Is " + scrip);
		    	        break;
		    	        
		    	    }
		    	}
 */

package KeatPro;

import java.awt.RenderingHints.Key;
import java.awt.Robot;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.io.IOException;
import java.net.URL;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.winium.DesktopOptions;
import org.openqa.selenium.winium.WiniumDriver;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

public class vitloop 
	{
	
		public static String[] getordno;
		public static double roundoff,cnvrtcaputil,placecnvrtcaputil,matchround;
		public static ExtentReports report1;
		public static ExtentTest log;
		public static WiniumDriver w;
		public static String caputilscrip;
		WebElement keatpro;
		
		 @Test
		 public void test() throws IOException, Exception
		 {
			  DesktopOptions options= new DesktopOptions();
			  options.setApplicationPath("D:\\KeatProX_UAT\\KeatProVX.exe");
			  
			  w=new WiniumDriver(new URL("http://localhost:9999"),options);
			  
			  WebDriverWait wait= new WebDriverWait (w,3000);
			  WebElement element=wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("1006")));
			  element.click();
			  
			  //WebElement window =  w.findElement(By.name("Login Details"));
			  
			  //Thread.sleep(15000);
			  
			  //login to keatpro
			  w.findElement(By.id("1006")).sendKeys("1y039");
			  w.findElement(By.id("1007")).sendKeys("login1");
			  w.findElementById("1014").sendKeys("1111");	
			  w.findElementByName("Login").click();
			  
			  WebDriverWait wait1= new WebDriverWait (w,3000);
			  WebElement element1=wait1.until(ExpectedConditions.visibilityOfElementLocated(By.name("Indices")));
			  //element1.click();				  
			  //keatpro = w.findElement(By.name("KEATProX Ver 2.2"));
			  
			  //Thread.sleep(35000);
			  
			  //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			  ///capture util
			  Robot robot = new Robot();
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_R);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_P);
			  Thread.sleep(500);
			  robot.keyRelease(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  
			  String scrip1 = "DLF~EQ";
//////////////////////////////////////////////////////			  
			  WebElement op = w.findElement(By.name("Net Position")).findElement(By.id("33302"));				   
			  //WebElement op = w.findElement(By.name("Net Position"));
			  List<WebElement> get = op.findElements(By.id("33302"));
			  System.out.println("get list szie "+get.size());
			  
			  if(get.equals(scrip1))
			  {
				  System.out.println("scrip found ");
			  }
			  else
			  {
				  System.out.println("scrip not found ");
			  }				  
			  
/////////////////////////////////////////////////////	
			  
			  
/*			  
			  for(int i=1;i<6;i++)
			  {
				  WebElement op = w.findElement(By.name("Net Position")).findElement(By.id("33302"));
				  WebElement abc = op.findElement(By.name("Row " + (i) + ""));
				  caputilscrip = abc.findElement(By.name("Row " + (i) + ", Column 1")).getText();
				  System.out.println("cap util "+caputilscrip);
				  if(scrip1.equals(caputilscrip))
				  {
					  String caputil = abc.findElement(By.name("Row " + (i) + ", Column 18")).getText();
					  System.out.println("cap util "+caputil);
					  cnvrtcaputil = Double.parseDouble(caputil);
					  break;
				  }
				  else
				  {
					  Double caputil = 0.0;
					  System.out.println("Scrip not available to capture position");
				  }				  
			  }	
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
			  report1 = new ExtentReports("D:\\Windowapp\\winapputil.html");
			  log = report1.startTest("Open Watchlist");  
			  				
			  ///select watchlist
			  String WL = "UTILISATION";
			  w.findElement(By.id("1012")).click();
			  w.findElement(By.name("UTILISATION")).click();	  
			  w.findElement(By.id("1")).click();
			  log.log(LogStatus.INFO, "Watchlist opened");
			  
			  //w.getWindowHandle().startsWith(scrip1);
			  keatpro = w.findElement(By.name("Watchlist - " + (WL) + "")).findElement(By.id("1000"));
			  //String abc = w.findElement(By.xpath("//*[contains(@Name, 'Row 1, Column 1')]")).getText();
			  //System.out.println("abc "+abc);
///////////////////////////////////////////////////////
			  if(abc.equals(scrip1))
			  {
				  keatpro.findElement(By.xpath("//*[contains(@Name, '"+scrip1+"')]")).click();
				  System.out.println("clicked on scrip ");
			  }
			  else
			  {
				  System.out.println("not clicked on scrip ");
			  }			
			  
			  ///select scrip			  
			  for(int i=1;i<10;i++)
			  {				  				  
				  //WebElement scrip = w.findElementByClassName("MDIClient");
				  //String getscrip = scrip.findElement(By.name("Row " + (i) + ", Column 1")).getText();
				  String getscrip = keatpro.findElement(By.name("Row " + (i) + ", Column 1")).getText();
				  //System.out.println("get scrip "+getscrip);
				  
				  if(scrip1.equals(getscrip))
				  {
					  keatpro.findElement(By.name("Row " + (i)+ ", Column 1")).click();
					  System.out.println("scrip got ");
					  log.log(LogStatus.INFO, "Scrip " +getscrip+ " got selected");
					  
					  ///place order
					  break;
				  }
				  else
				  {
					  log.log(LogStatus.ERROR, "Scrip not found");
					  System.out.println("scrip not found ");
				  }	    					 	    	

			  }			  			  
			  			  
			  //Runtime.getRuntime().exec("D:\\Windowapp\\White-master\\tools\\UISpy.exe");
			  
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
			  w.findElement(By.id("1350")).sendKeys("2");
			  robot.keyPress(KeyEvent.VK_TAB);
			  Thread.sleep(500);
			  String getamt = w.findElement(By.id("1409")).getText();
			  System.out.println("getamount "+getamt);
			  Double getamt1 = Double.parseDouble(getamt);
			  Double calc = getamt1 - 20.0;	
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
			  			  
			  ///capture util
			  robot.keyPress(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_R);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_P);
			  Thread.sleep(500);
			  robot.keyRelease(KeyEvent.VK_ALT);
			  Thread.sleep(500);			  
			  
			  for(int i=1;i<6;i++)
			  {
				  WebElement op = w.findElement(By.name("Net Position")).findElement(By.id("33302"));				  				  
				  String placecaputilscrip = op.findElement(By.name("Row " + (i) + ", Column 1")).getText();
				  System.out.println("place cap util scrip "+placecaputilscrip);
				  if(scrip1.equals(placecaputilscrip))
				  {					  
					  //op.findElement(By.name("Row " + (i) + ", Column 1")).click();
					  
					  ///refresh scrip
					  robot.mousePress(InputEvent.BUTTON3_DOWN_MASK);
					  Thread.sleep(500);
					  robot.mouseRelease(InputEvent.BUTTON3_DOWN_MASK);
					  Thread.sleep(500);
					  robot.keyPress(KeyEvent.VK_R);
					  Thread.sleep(500);
					  
					  WebElement opp = w.findElement(By.name("Net Position")).findElement(By.id("33302"));
					  WebElement abc = opp.findElement(By.name("Row " + (i) + ""));
					  
					  String placecaputil = abc.findElement(By.name("Row " + (i) + ", Column 18")).getText();
					  System.out.println("cap new util "+placecaputil);
					  placecnvrtcaputil = Double.parseDouble(placecaputil);
					  System.out.println("cnvrt cap util again "+cnvrtcaputil);
					  Double match = placecnvrtcaputil - cnvrtcaputil;
					  System.out.println("minus " +match);
					  matchround = Math.round(match*100.0)/100.0;
					  System.out.println("round minus " +matchround);						
					  log.log(LogStatus.INFO, "Margin utilised is : " +matchround);
					  break;
				  }
				  else
				  {					  
					  System.out.println("Scrip not available to capture position");
				  }
				  
			  }
			  //open order report
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_R);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_O);
			  Thread.sleep(500);
			  robot.keyRelease(KeyEvent.VK_ALT);
			  Thread.sleep(500);			  
			  
			  keatpro =  w.findElement(By.name("Order Report [All]"));
		      		        
			  ///calculate util for place			  
			  for(int i=1;i<10;i++)
			  {				  				  
				  WebElement orderrpt = w.findElement(By.name("Order Report [All]")).findElement(By.id("1003"));
				  String capordno = orderrpt.findElement(By.name("Row " + (i) + ", Column 4")).getText();
				  System.out.println("get scrip "+capordno);
				  
				  //WebElement menuItem = keatpro.findElement(By.name("GridControl"));//.findElement(By.name("Row " + (i) + ""));
			      //String capordno = menuItem.findElement(By.name("Row " + (i) + ", Column 4")).getText();
			      
				  if(getordno[1].equals(capordno))
				  {
					  orderrpt.findElement(By.name("Row " + (i)+ ", Column 4")).click();					  
					  String getqty = orderrpt.findElement(By.name("Row " + (i)+ ", Column 7")).getText();
					  int qty = Integer.parseInt(getqty);
					  Double calcutil = (calc * qty)/7.75;
					  roundoff = (Math.round(calcutil*100.0)/100.0);
					  System.out.println("calc util "+roundoff);				  
					  log.log(LogStatus.INFO, "Required Margin is : "+roundoff);
					  
					  ////calc margin utilised
					  
					  
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
					  
					  ///modify order
					  
					  String getval = orderrpt.findElement(By.name("Row " + (i)+ ", Column 4")).getText();
					  System.out.println("value "+getval);
					  
					  robot.keyPress(KeyEvent.VK_F5);
					  Thread.sleep(500);
					  
					  String getstat = orderrpt.findElement(By.name("Row " + (i)+ ", Column 11")).getText();
					  System.out.println("status "+getstat);
					  
					  if(getval.equals(getordno[1]) && getstat.equals("OPN"))
					  {
						  orderrpt.findElement(By.name("Row " + (i)+ "")).click();
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
						  w.findElement(By.id("6")).click();
						  Thread.sleep(1000);
						  w.findElement(By.id("1")).click();
						  Thread.sleep(500);
						  w.findElement(By.id("2")).click();	///cancel modify dialog box		
						  log.log(LogStatus.INFO, "Order no " +getordno[1]+ " modified successfully");
						  
						  ///cancel order
						  orderrpt.findElement(By.name("Row " + (i)+ "")).click();
						  Thread.sleep(500);
						  robot.mousePress(InputEvent.BUTTON3_DOWN_MASK);
						  Thread.sleep(500);
						  robot.mouseRelease(InputEvent.BUTTON3_DOWN_MASK);
						  Thread.sleep(500);
						  robot.keyPress(KeyEvent.VK_C);
						  Thread.sleep(500);
						  w.findElement(By.id("1360")).click();
						  w.findElement(By.id("6")).click();
						  Thread.sleep(1000);
						  w.findElement(By.id("1")).click();
						  Thread.sleep(500);
						  w.findElement(By.id("2")).click();	///cancel cancel dialog box
						  log.log(LogStatus.INFO, "Order no " +getordno[1]+ " cancelled successfully");
						  
					  }
					  else
					  {
						  log.log(LogStatus.INFO, "Order is " +getstat+ " so cannot modify and cancel.");
						  System.out.println("Order is " +getstat+ " so cannot modify and cancel. ");
					  }
					  
					  break;
				  }
				  else
				  {
					  System.out.println("order no not found ");
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
	*/	  
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

