package KeatDealer;

import java.awt.Robot;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.LogStatus;

import Libraries.functions;
import Libraries.variables;
import Main.beforesuitedealer;

public class TC4LoadFNOSMP extends beforesuitedealer
	{	
		public static String[] getordno,remdot;
		public static double cnvrtcaputil,placecnvrtcaputil,matchround;
		public static String caputilscrip,caputil,getmsg,capordno,getxchnge;
		public static String ME,capME,capqty,rghclient,otherclient;
		WebElement keatpro,op;		
		public static int i,j,xcelrc,m;
		Robot robot;
		public static int k=1;
		public static float SLOprice;
		public static double chngestrnggetamt;
		
		variables vr = new variables();
		functions fn = new functions();
		/*
		public static void main(String[] args) 
		  {
			  TestListenerAdapter tla = new TestListenerAdapter();
			    TestNG testng = new TestNG();
			     Class[] classes = new Class[]
			     {
			    	 TC2Utilisation.class     
			     };
			     testng.setTestClasses(classes);
			     testng.addListener(tla);
			     testng.run();		  	  
		  	}
		*/
		
		
		public void openWL() throws Exception
		{				  
			FileInputStream f = new FileInputStream(srcfile);
			wb = new HSSFWorkbook(f);		
			vr.s = wb.getSheet("PLFNO");
			 
			robot = new Robot();	
			
			///open watchlist
			Thread.sleep(500);
			robot.keyPress(KeyEvent.VK_ALT);
			Thread.sleep(500);
			robot.keyPress(KeyEvent.VK_F);
			Thread.sleep(500);
			robot.keyPress(KeyEvent.VK_O);
			Thread.sleep(500);
			robot.keyRelease(KeyEvent.VK_ALT);
			  
			///select watchlist
			String WL = vr.s.getRow(1).getCell(5).getStringCellValue();
			w.findElement(By.id("1010")).click();
			WebElement openwltable = w.findElement(By.name("WatchList")).findElement(By.id("1010"));
			openwltable.findElement(By.xpath("//*[contains(@Name, '"+WL+"')]")).click();
			
			String WL1 = vr.s.getRow(1).getCell(6).getStringCellValue();
			w.findElement(By.id("1013")).click();
			WebElement openwltable1 = w.findElement(By.name("WatchList")).findElement(By.id("1013"));
			openwltable1.findElement(By.xpath("//*[contains(@Name, '"+WL1+"')]")).click();
			
			w.findElement(By.id("1")).click();			
			f.close();
			
			rghclient = vr.s.getRow(3).getCell(5).getStringCellValue();
			otherclient = vr.s.getRow(2).getCell(5).getStringCellValue();
		}
		
		public void close() throws Exception
		{
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_F);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_C);
			  Thread.sleep(500);
			  robot.keyRelease(KeyEvent.VK_ALT);
			  Thread.sleep(500);
		}					
		
		public void write2excel() throws Exception
		{						
			FileInputStream f = new FileInputStream(srcfile);
			wb = new HSSFWorkbook(f);		
			vr.s = wb.getSheet("PLFNO");
			
			xcelrc = vr.s.getLastRowNum() - vr.s.getFirstRowNum();
		  	System.out.println("rc "+xcelrc);		  	

			WebElement wltable = w.findElement(By.id("65280")).findElement(By.id("1000"));
			wltable.findElement(By.name("Row " + (k)+ ", Column 1")).click();
			
			log = report1.startTest("Place Order");
			log.log(LogStatus.INFO, "Watchlist opened");	
			
			  for(i=5;i<=xcelrc;)
			  {
				  Row row = vr.s.getRow(i);	
						  						  
						  for(j=3;j<row.getLastCellNum();)
						  {			
							///select action							  
							  vr.action = vr.s.getRow(i).getCell(j).getStringCellValue();

							  if(vr.action.equals("BUY"))							  
							  {
								  robot.keyPress(KeyEvent.VK_CONTROL);
								  robot.keyPress(KeyEvent.VK_F1);
								  robot.keyRelease(KeyEvent.VK_CONTROL);
								  robot.keyRelease(KeyEvent.VK_F1);
							  }
							  if(vr.action.equals("SELL"))							  
							  {
								  robot.keyPress(KeyEvent.VK_CONTROL);
								  robot.keyPress(KeyEvent.VK_F2);
								  robot.keyRelease(KeyEvent.VK_CONTROL);
								  robot.keyRelease(KeyEvent.VK_F2);
							  }
							  
							  ///enter client code
							  j=j+1;
							  vr.clientcode = vr.s.getRow(i).getCell(j).getStringCellValue();
							  Thread.sleep(3000);
							  w.findElement(By.id("1018")).sendKeys(vr.clientcode);
							  
							  robot.keyPress(KeyEvent.VK_TAB);
							  //robot.keyPress(KeyEvent.VK_DELETE);	
							  
							  ///enter qty
							  j=j+1;
							  capqty = w.findElement(By.id("1400")).getText();
							  //vr.qty = (int) vr.s.getRow(i).getCell(j).getNumericCellValue();
							  //w.findElement(By.id("1400")).sendKeys(String.valueOf(capqty));
							  log.log(LogStatus.INFO, "Entered quantity "+capqty);
							 							  
							  robot.keyPress(KeyEvent.VK_TAB);
							  vr.getamt = w.findElement(By.id("1404")).getText();
							  System.out.println("getamount "+vr.getamt);							  
				  					  
							  ///get price perc
							  j=j+1;
							  vr.perc = (float) vr.s.getRow(i).getCell(j).getNumericCellValue();
							  vr.getamt1 = (Double.parseDouble(vr.getamt)*vr.perc)/100;
							  chngestrnggetamt = Double.parseDouble(vr.getamt);
							  
							  if(vr.action.equals("BUY"))
								{
									vr.total = chngestrnggetamt-vr.getamt1;
								}
							  else
								{
									vr.total = chngestrnggetamt+vr.getamt1;
								}
							  vr.roundoff = (double) (Math.round(vr.total));
							  log.log(LogStatus.INFO, "Actual price is " +vr.getamt);
							  								  								  
							  robot.keyPress(KeyEvent.VK_DELETE);
							  
							  if(vr.roundoff<=0)
							  {
								  vr.roundoff=0.0;
							  }									  
							  if(vr.perc==0)
							  {
								  vr.roundoff=0.0;
							  }
							  
							  w.findElement(By.id("1404")).sendKeys(String.valueOf(vr.roundoff));								  							  
							  log.log(LogStatus.INFO, "Modified price is " +vr.roundoff);	

							  j=j+1;
							  //vr.rw = vr.s.getRow(i);
							  //vr.cell = vr.rw.createCell(j);
							  //vr.cell.setCellValue(vr.getamt);
							  
							  j=j+1;
							  //vr.rw = vr.s.getRow(i);
							  //vr.cell = vr.rw.createCell(j);
							  //vr.cell.setCellValue(vr.roundoff);
							  
							  ///select order mark
							  j=j+1;
							  String ordermrk = vr.s.getRow(i).getCell(j).getStringCellValue();
							  w.findElement(By.id("1014")).click();
							  log.log(LogStatus.INFO, "Select Order Mark");
							  if(ordermrk.equals("Normal"))
							  {
								  w.findElement(By.name("Normal")).click();
								  if((vr.roundoff!=0))// || (vr.roundoff==0) )
								  {
									  w.findElement(By.id("1342")).click();/////new buy order button
									  Thread.sleep(1500);
									  w.findElement(By.id("1")).click();/////success button
									  Thread.sleep(1500);
									  
									  w.findElement(By.id("1018")).sendKeys(otherclient);
									  if(vr.action.equals("SELL"))							  
									  {
										  robot.keyPress(KeyEvent.VK_CONTROL);
										  robot.keyPress(KeyEvent.VK_F1);
										  robot.keyRelease(KeyEvent.VK_CONTROL);
										  robot.keyRelease(KeyEvent.VK_F1);
									  }
									  else if(vr.action.equals("BUY"))							  
									  {
										  robot.keyPress(KeyEvent.VK_CONTROL);
										  robot.keyPress(KeyEvent.VK_F2);
										  robot.keyRelease(KeyEvent.VK_CONTROL);
										  robot.keyRelease(KeyEvent.VK_F2);
									  }
									  w.findElement(By.id("1400")).sendKeys(String.valueOf(capqty));
									  w.findElement(By.id("1342")).click();/////new buy order button
									  Thread.sleep(1500);
									  w.findElement(By.id("1")).click();/////success button
									  Thread.sleep(1500);
									  w.findElement(By.id("2")).click();
//////////////////////////////////////////////////////////////////////////////////
									robot.keyPress(KeyEvent.VK_DELETE);
									robot.keyPress(KeyEvent.VK_ENTER);
									robot.keyPress(KeyEvent.VK_PAGE_DOWN);
									Thread.sleep(2000);
//////////////////////////////////////////////////////////////////////////////////
									  break;
								  }
								  else if((vr.roundoff==0))
								  {
									  if(vr.perc !=0)
									  {		
										  w.findElement(By.id("1342")).click();/////new buy order button
										  Thread.sleep(2000);
										  log.log(LogStatus.INFO, "Click on New "+vr.action+ " Order ");
										  			
										  robot.keyPress(KeyEvent.VK_ENTER);
										  Thread.sleep(2000);
										  
										  w.findElement(By.id("1018")).sendKeys(otherclient);
										  log.log(LogStatus.INFO, "Enter other Client " +otherclient);	

										  if(vr.action.equals("SELL"))							  
										  {
											  robot.keyPress(KeyEvent.VK_CONTROL);
											  robot.keyPress(KeyEvent.VK_F1);
											  robot.keyRelease(KeyEvent.VK_CONTROL);
											  robot.keyRelease(KeyEvent.VK_F1);
										  }
										  else if(vr.action.equals("BUY"))							  
										  {
											  robot.keyPress(KeyEvent.VK_CONTROL);
											  robot.keyPress(KeyEvent.VK_F2);
											  robot.keyRelease(KeyEvent.VK_CONTROL);
											  robot.keyRelease(KeyEvent.VK_F2);
										  }
										  log.log(LogStatus.INFO, "Selected action is " +vr.action);	
										  w.findElement(By.id("1400")).sendKeys(capqty);
										  log.log(LogStatus.INFO, "Entered quantity " +vr.qty);	

										  w.findElement(By.id("1404")).sendKeys("0.00"); ///for trade at mkt
										  log.log(LogStatus.INFO, "Place order at market price ");	

										  w.findElement(By.id("1342")).click();/////new buy order button
										  Thread.sleep(2000);
										  w.findElement(By.id("1")).click();/////success button
										  Thread.sleep(2000);
										  w.findElement(By.id("2")).click();
//////////////////////////////////////////////////////////////////////////////////
										robot.keyPress(KeyEvent.VK_DELETE);
										robot.keyPress(KeyEvent.VK_ENTER);
										robot.keyPress(KeyEvent.VK_PAGE_DOWN);
										Thread.sleep(2000);
//////////////////////////////////////////////////////////////////////////////////
										  break;
									  }
									  else
									  {
										  w.findElement(By.id("1342")).click();/////new buy order button
										  Thread.sleep(2000);
										  w.findElement(By.id("1")).click();/////success button
										  Thread.sleep(2000);
										  w.findElement(By.id("2")).click();
										  log.log(LogStatus.INFO, "Place order at market price ");
										  
										  //trade order
										  if(vr.action.equals("SELL"))							  
										  {
											  robot.keyPress(KeyEvent.VK_CONTROL);
											  robot.keyPress(KeyEvent.VK_F1);
											  robot.keyRelease(KeyEvent.VK_CONTROL);
											  robot.keyRelease(KeyEvent.VK_F1);
										  }
										  else if(vr.action.equals("BUY"))							  
										  {
											  robot.keyPress(KeyEvent.VK_CONTROL);
											  robot.keyPress(KeyEvent.VK_F2);
											  robot.keyRelease(KeyEvent.VK_CONTROL);
											  robot.keyRelease(KeyEvent.VK_F2);
										  }
										  log.log(LogStatus.INFO, "Selected action is " +vr.action);
										  w.findElement(By.id("1018")).sendKeys(otherclient);
										  log.log(LogStatus.INFO, "Enter other Client " +otherclient);	

										  w.findElement(By.id("1400")).sendKeys(capqty);
										  log.log(LogStatus.INFO, "Entered quantity " +vr.qty);	

										  w.findElement(By.id("1404")).sendKeys("0.00"); ///for trade at mkt
										  log.log(LogStatus.INFO, "Place order at market price ");	

										  w.findElement(By.id("1342")).click();/////new buy order button
										  Thread.sleep(2000);
										  w.findElement(By.id("1")).click();/////success button
										  Thread.sleep(2000);
										  w.findElement(By.id("2")).click();
//////////////////////////////////////////////////////////////////////////////////
										robot.keyPress(KeyEvent.VK_DELETE);
										robot.keyPress(KeyEvent.VK_ENTER);
										robot.keyPress(KeyEvent.VK_PAGE_DOWN);
										Thread.sleep(2000);
//////////////////////////////////////////////////////////////////////////////////
										  break;
									  }
								  }
							  }
							  else if(ordermrk.equals("Margin Trading"))
							  {
								  w.findElement(By.name("Margin Trading")).click();	
							  }
							  else if(ordermrk.equals("Super Multiple Plus"))
							  {
								  if(vr.roundoff!=0)
								  {
									  w.findElement(By.name("Super Multiple Plus")).click();
									  j=j+1;
									  SLOprice = (float) vr.s.getRow(i).getCell(j).getNumericCellValue();
									  vr.HLprice = (vr.roundoff*SLOprice)/100;
									  System.out.println("hlprice "+vr.HLprice);
									  if(vr.action.equals("BUY"))
									  {
										vr.calHLprice = vr.roundoff-vr.HLprice;
									  }
									  else if(vr.action.equals("SELL"))
									  {
										vr.calHLprice = vr.roundoff+vr.HLprice;
									  }
									  vr.roundoff1 = (double) (Math.round(vr.calHLprice));
									  w.findElement(By.id("1405")).sendKeys(String.valueOf(vr.roundoff1));	
									  w.findElement(By.id("1342")).click();/////new buy order button
									  Thread.sleep(2000);
									  w.findElement(By.id("1")).click();/////success button
									  Thread.sleep(2000);
									  w.findElement(By.id("2")).click();												  
//////////////////////////////////////////////////////////////////////////////////
										robot.keyPress(KeyEvent.VK_DELETE);
										robot.keyPress(KeyEvent.VK_ENTER);
										robot.keyPress(KeyEvent.VK_PAGE_DOWN);
										Thread.sleep(2000);
//////////////////////////////////////////////////////////////////////////////////
										break;
								  }
								  
								  
								  else if(vr.roundoff==0)
								  {
									  for(int a=1;a<=1;a++) 
									  {
										  w.findElement(By.id("2")).click();
										  
										  if(vr.action.equals("SELL"))							  
										  {
											  Thread.sleep(2000);
											  robot.keyPress(KeyEvent.VK_F1);
											  robot.keyRelease(KeyEvent.VK_F1);
										  }
										  else if(vr.action.equals("BUY"))							  
										  {
											  Thread.sleep(2000);
											  robot.keyPress(KeyEvent.VK_F2);
											  robot.keyRelease(KeyEvent.VK_F2);
										  }
										  w.findElement(By.id("1018")).sendKeys(rghclient);
										  w.findElement(By.id("1400")).sendKeys(capqty);
										  w.findElement(By.id("1014")).click();
										  w.findElement(By.name("Normal")).click();							
										  w.findElement(By.id("1342")).click();/////new buy order button
										  Thread.sleep(1500);
										  w.findElement(By.id("1")).click();/////success button
										  Thread.sleep(1500);
										  w.findElement(By.id("2")).click();
										  
										  //place sm order
										  if(vr.action.equals("BUY"))							  
										  {
											  robot.keyPress(KeyEvent.VK_CONTROL);
											  robot.keyPress(KeyEvent.VK_F1);
											  robot.keyRelease(KeyEvent.VK_CONTROL);
											  robot.keyRelease(KeyEvent.VK_F1);
										  }
										  else if(vr.action.equals("SELL"))							  
										  {
											  robot.keyPress(KeyEvent.VK_CONTROL);
											  robot.keyPress(KeyEvent.VK_F2);
											  robot.keyRelease(KeyEvent.VK_CONTROL);
											  robot.keyRelease(KeyEvent.VK_F2);
										  }									  
										  w.findElement(By.id("1018")).sendKeys(vr.clientcode);
										  w.findElement(By.id("1400")).sendKeys(capqty);
										  robot.keyPress(KeyEvent.VK_TAB);
										  vr.getamt = w.findElement(By.id("1404")).getText();
										  System.out.println("perc "+vr.perc);
										  vr.getamt1 = (Double.parseDouble(vr.getamt)*vr.perc)/100;
										  chngestrnggetamt = Double.parseDouble(vr.getamt);
										  
										  if(vr.action.equals("BUY"))
											{
												vr.total = chngestrnggetamt-vr.getamt1;
											}
										  else
											{
												vr.total = chngestrnggetamt+vr.getamt1;
											}
										  vr.roundoff = (double) (Math.round(vr.total));
										  if(vr.roundoff<=0)
										  {
											  vr.roundoff=0.0;
										  }									  
										  if(vr.perc==0)
										  {
											  vr.roundoff=0.0;
										  }
										  w.findElement(By.id("1014")).click();
										  w.findElement(By.name("Super Multiple Plus")).click();
										  j=j+1;
										  SLOprice = (float) vr.s.getRow(i).getCell(j).getNumericCellValue();
										  vr.HLprice = (vr.roundoff*SLOprice)/100;
										  System.out.println("hlprice "+vr.HLprice);
										  if(vr.action.equals("BUY"))
										  {
											vr.calHLprice = vr.roundoff-vr.HLprice;
										  }
										  else if(vr.action.equals("SELL"))
										  {
											vr.calHLprice = vr.roundoff+vr.HLprice;
										  }
										  vr.roundoff1 = (double) (Math.round(vr.calHLprice));
										  w.findElement(By.id("1405")).sendKeys(String.valueOf(vr.roundoff1));											  
										  w.findElement(By.id("1342")).click();/////new buy order button
										  Thread.sleep(2000);
										  w.findElement(By.id("1")).click();/////success button
										  Thread.sleep(2000);
										  w.findElement(By.id("2")).click();	
										  
										////trade sm order				  
										  if(vr.action.equals("SELL"))							  
										  {
											  robot.keyPress(KeyEvent.VK_CONTROL);
											  robot.keyPress(KeyEvent.VK_F1);
											  robot.keyRelease(KeyEvent.VK_CONTROL);
											  robot.keyRelease(KeyEvent.VK_F1);
										  }
										  else if(vr.action.equals("BUY"))							  
										  {
											  robot.keyPress(KeyEvent.VK_CONTROL);
											  robot.keyPress(KeyEvent.VK_F2);
											  robot.keyRelease(KeyEvent.VK_CONTROL);
											  robot.keyRelease(KeyEvent.VK_F2);
										  }
										  w.findElement(By.id("1018")).sendKeys(otherclient);
										  w.findElement(By.id("1400")).sendKeys(capqty);
										  w.findElement(By.id("1014")).click();
										  w.findElement(By.name("Super Multiple Plus")).click();
										  vr.HLprice = (vr.roundoff*SLOprice)/100;
										  System.out.println("hlprice "+vr.HLprice);
										  if(vr.action.equals("BUY"))
										  {
											vr.calHLprice = vr.roundoff+vr.HLprice; ///dis change for trading purpose as action is taking vice versa
										  }
										  else if(vr.action.equals("SELL"))
										  {
											vr.calHLprice = vr.roundoff-vr.HLprice; ///dis change for trading purpose as action is taking vice versa
										  }
										  vr.roundoff1 = (double) (Math.round(vr.calHLprice));
										  w.findElement(By.id("1405")).sendKeys(String.valueOf(vr.roundoff1));
										  System.out.println("roundoff "+vr.roundoff+ " roundoff1" +vr.roundoff1);
										  
										  w.findElement(By.id("1342")).click();/////new buy order button
										  Thread.sleep(2000);
										  robot.keyPress(KeyEvent.VK_ENTER);
										  Thread.sleep(2000);
										  w.findElement(By.id("2")).click();												  
										  Thread.sleep(2000);
										  log.log(LogStatus.INFO, "Trade super multiple plus order ");
										  
//////////////////////////////////////////////////////////////////////////////////
										robot.keyPress(KeyEvent.VK_DELETE);
										robot.keyPress(KeyEvent.VK_ENTER);
										robot.keyPress(KeyEvent.VK_PAGE_DOWN);
										Thread.sleep(2000);
//////////////////////////////////////////////////////////////////////////////////
										  break;
									}
									  break;
								  }
							  }
							  else if(ordermrk.equals("Buy Delivery"))
							  {
								  w.findElement(By.name("Buy Delivery")).click();
							  }
							  else if(ordermrk.equals("Super Multiple Option"))
							  {
								  w.findElement(By.name("Super Multiple Option")).click();
							  }
							  
							  w.findElement(By.id("1342")).click();/////new buy order button
							  Thread.sleep(2500);
							  log.log(LogStatus.INFO, "Click on New "+vr.action+ " Order ");			
							  		
							  //success enter
							  w.findElement(By.id("1")).click();/////new buy order button
							  Thread.sleep(2000);
							  
							  //cancel								  
							  w.findElement(By.id("2")).click();
							  log.log(LogStatus.INFO, "Click on Cancel ");	
							  j=3;
							  break;	
				  }	
						  
						i++;						
			  }
			  FileOutputStream Of = new FileOutputStream(srcfile);
				
				wb.write(Of);
				Of.close();
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
			  log.log(LogStatus.INFO, "Order Report refreshed ");
			  WebElement rpttable = w.findElement(By.name("Order Report [All]"));
			  rpttable.findElement(By.id("1018")).sendKeys(vr.clientcode);
			  robot.keyPress(KeyEvent.VK_F5);
			  Thread.sleep(500);
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
			  limittable.findElement(By.id("1018")).sendKeys(vr.clientcode);
			  robot.keyPress(KeyEvent.VK_F5);
			  Thread.sleep(500);			  
		}
		
		public void exit() throws Exception
		{
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
		
		 @Test
		 public void placemodcanorder() throws IOException, Exception
		 {		  		  	
			 			  			  
			  ///report creation
			  report1 = new ExtentReports("D:\\Windowapp\\positionload.html");
			  log = report1.startTest("Login Details");
			  
			  ////call login function
			  Login();
			  
			  robot = new Robot();
			  
			  ////open watchlist
			  openWL();
			  
			  write2excel();
			  ordrptf5();
			  limitpgf5();
			  exit();
			  			  
		 	}
		 
		 
		 @Test(priority=4001)
			public void nouse()
			  {
				  report1.endTest(log);
				  report1.flush();
				  //w.get("D:\\Windowapp\\positionload.html");				  
				  //w.get(srcfile);
			  }
		 
		}

/*
var peer = FrameworkElementAutomationPeer.FromElement(element);

if (peer == null)
{
    throw new Exception();
}

var s = peer.GetPattern(PatternInterface.Scroll) as IScrollProvider;

if (s == null)
{
    throw new Exception();
}

if (!s.VerticallyScrollable)
{
    throw new Exception();
}

s.Scroll(ScrollAmount.NoAmount, ScrollAmount.SmallIncrement)
*/