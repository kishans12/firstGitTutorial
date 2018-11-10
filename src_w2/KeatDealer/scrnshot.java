package KeatDealer;

import java.awt.Robot;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.TestListenerAdapter;
import org.testng.TestNG;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.LogStatus;

import Libraries.functions;
import Libraries.variables;
import Main.beforesuitedealer;

public class scrnshot extends beforesuitedealer
	{	
		public static String[] getordno,remdot,gettsloordno,tsloremdot,getsmordno,smremdot;
		public static double cnvrtcaputil,placecnvrtcaputil,matchround,chngegetprice,chngestrnggetamt;
		public static String caputilscrip,caputil,getmsg,capordno,getxchnge,TSLOchk,tsloordno,smordno;
		public static String ME,capME,otherclient,capscrip,delmark,getstat,nettrdval,getqty;
		WebElement keatpro,op;		
		public static int i,j,z,xcelrc,m,addqty,bpprice;
		Robot robot;
		public static float SLOprice,spread;
		public static String gettrigpr,ordermrk,roughclient;
		public static double totbp;
		public static String bpr,sqty,bqty,spr,bord,sord;
		public static String storenormordno,storesmordno,storetsloordno;
		
		variables vr = new variables();
		functions fn = new functions();	
		/*	
		public static void main(String[] args) 
		  {
			  TestListenerAdapter tla = new TestListenerAdapter();
			    TestNG testng = new TestNG();
			     Class[] classes = new Class[]
			     {
			    	 scrnshot.class
			     };
			     testng.setTestClasses(classes);
			     testng.addListener(tla);
			     testng.run();		  	  
		  	}
		*/
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
		
		public void tslorpt() throws Exception
		{		  
			  robot.keyPress(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_R);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_L);
			  Thread.sleep(1500);	
			  robot.keyRelease(KeyEvent.VK_ALT);
			  Thread.sleep(1500);
			  WebElement rpttable = w.findElement(By.name("TSLO Order Report"));
			  rpttable.findElement(By.id("1018")).sendKeys(vr.clientcode);
			  robot.keyPress(KeyEvent.VK_ENTER);
			  Thread.sleep(500);
			  robot.keyRelease(KeyEvent.VK_ENTER);
			  Thread.sleep(500);
			  log.log(LogStatus.INFO, "TSLO Order Report ");
		}
		
		public void orderrpt() throws Exception
		{		  
			  robot.keyPress(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_R);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_O);
			  Thread.sleep(1500);	
			  robot.keyRelease(KeyEvent.VK_ALT);
			  Thread.sleep(1500);
			  WebElement rpttable = w.findElement(By.name("Order Report [All]"));
			  rpttable.findElement(By.id("1018")).sendKeys(vr.clientcode);
			  robot.keyPress(KeyEvent.VK_F5);
			  Thread.sleep(500);
			  log.log(LogStatus.INFO, "Order Report ");
		}		
				
		public void openposition() throws Exception
		{		  		  
			  robot.keyPress(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_R);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_P);
			  Thread.sleep(500);
			  robot.keyRelease(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  WebElement limittable = w.findElement(By.name("Net Position"));
			  limittable.findElement(By.id("1018")).sendKeys(vr.clientcode);
			  robot.keyPress(KeyEvent.VK_F5);
			  Thread.sleep(500);	
			  log.log(LogStatus.INFO, "Open position report ");		  
		}
		
		public void todaysposition() throws Exception
		{		  		  
			  robot.keyPress(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_R);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_D);
			  Thread.sleep(500);
			  robot.keyRelease(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  WebElement limittable = w.findElement(By.name("Net Position"));
			  limittable.findElement(By.id("1018")).sendKeys(vr.clientcode);
			  robot.keyPress(KeyEvent.VK_F5);
			  Thread.sleep(500);	
			  log.log(LogStatus.INFO, "Today's position ");
		}
		
		public void traderpt() throws Exception
		{		  
			  robot.keyPress(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_R);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_T);
			  Thread.sleep(1500);	
			  robot.keyRelease(KeyEvent.VK_ALT);
			  Thread.sleep(1500);
			  WebElement rpttable = w.findElement(By.name("Trade Report [All]"));
			  rpttable.findElement(By.id("1018")).sendKeys(vr.clientcode);
			  robot.keyPress(KeyEvent.VK_F5);
			  Thread.sleep(500);
			  log.log(LogStatus.INFO, "Trade Report ");
		}
		
		public void limitspg() throws Exception
		{		  		  
			  robot.keyPress(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_R);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_P);
			  Thread.sleep(500);
			  robot.keyRelease(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_SHIFT);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_L);
			  Thread.sleep(500);
			  robot.keyRelease(KeyEvent.VK_SHIFT);
			  Thread.sleep(500);
			  WebElement limittable = w.findElement(By.name("Net Position"));
			  limittable.findElement(By.id("1018")).sendKeys(vr.clientcode);
			  robot.keyPress(KeyEvent.VK_F5);
			  Thread.sleep(500);	
			  log.log(LogStatus.INFO, "Limit's page ");
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
			
		public void openWL() throws Exception
		{				  
			FileInputStream f = new FileInputStream(srcfile);
			wb = new HSSFWorkbook(f);		
			vr.s = wb.getSheet("Mock");
			
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
			
			log = report1.startTest("Open Watchlist");
			///select watchlist
			roughclient = vr.s.getRow(2).getCell(5).getStringCellValue();
			System.out.println("rough client "+roughclient);
			String WL = vr.s.getRow(1).getCell(5).getStringCellValue();
			w.findElement(By.id("1010")).click();
			WebElement openwltable = w.findElement(By.name("WatchList")).findElement(By.id("1010"));
			openwltable.findElement(By.xpath("//*[contains(@Name, '"+WL+"')]")).click();
			log.log(LogStatus.INFO, "Select " +WL);

			String WL1 = vr.s.getRow(1).getCell(6).getStringCellValue();
			w.findElement(By.id("1013")).click();
			WebElement openwltable1 = w.findElement(By.name("WatchList")).findElement(By.id("1013"));
			openwltable1.findElement(By.xpath("//*[contains(@Name, '"+WL1+"')]")).click();
			log.log(LogStatus.INFO, "Select "+WL1);
			
			w.findElement(By.id("1")).click();
			f.close();
			log.log(LogStatus.INFO, "Rough Client " +roughclient);			
		}
		
		public void placeorder() throws Exception
		{			
			FileInputStream f = new FileInputStream(srcfile);
			wb = new HSSFWorkbook(f);		
			vr.s = wb.getSheet("Mock");	
			
			WebElement wltable = w.findElement(By.id("65280")).findElement(By.id("1000"));
			WebElement colcount = wltable.findElement(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));		
			List < WebElement > rc1 = wltable.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			System.out.println("row count "+ rc1.size());
			
			xcelrc = vr.s.getLastRowNum() - vr.s.getFirstRowNum();
		  	System.out.println("rc "+xcelrc);
			  
			  for(i=5;i<xcelrc+1;i++)
			  {
				  Row row = vr.s.getRow(i);
				  
				  String flag = vr.s.getRow(i).getCell(0).getStringCellValue();
				  
				  if(flag.equals("Y"))
				  {
					  for(j=1;j<row.getLastCellNum();)
					  {		
						  vr.scrip = vr.s.getRow(i).getCell(j).getStringCellValue();
						  
						  ///excel for BSE
						  j=j+1;
						  		
						  ///select action
						  j=j+1;
						  vr.action = vr.s.getRow(i).getCell(j).getStringCellValue();
						  
						  ///enter client code
						  j=j+1;
						  vr.clientcode = vr.s.getRow(i).getCell(j).getStringCellValue();
						  
						  ///enter qty
						  j=j+1;
						  vr.qty = (int) vr.s.getRow(i).getCell(j).getNumericCellValue();
						  
						  ///get price perc
						  j=j+1;
						  vr.perc = (float) vr.s.getRow(i).getCell(j).getNumericCellValue();
						  
						  ///select order mark
						  j=j+1;
						  ordermrk = vr.s.getRow(i).getCell(j).getStringCellValue();

						  j=j+1;
						  
						  j=j+1;
						  otherclient = vr.s.getRow(i).getCell(j).getStringCellValue();
						  
						  j=j+1;
						  addqty = (int) vr.s.getRow(i).getCell(j).getNumericCellValue();
						  
						  j=j+1;
						  
						  j=j+1;
						  
						  j=j+1;
						  
						  j=j+1;						  
						  
						  for(int k=1;k<=rc1.size();k++)
						  {
							  vr.getscrip = wltable.findElement(By.name("Row " + (k) + ", Column 1")).getText();
							  System.out.println("get scrip " +vr.getscrip);
							  wltable.findElement(By.name("Row " + (k)+ ", Column 1")).click();
							  getxchnge = wltable.findElement(By.name("Row " + (k)+ ", Column 2")).getText();
							  System.out.println("get exchnage "+getxchnge);
							  vr.bprice = wltable.findElement(By.name("Row " + (k)+ ", Column 5")).getText();
							  vr.sprice = wltable.findElement(By.name("Row " + (k)+ ", Column 7")).getText();

							  vr.getprice = wltable.findElement(By.name("Row " + (k)+ ", Column 9")).getText();
							  System.out.println("get price "+vr.getprice);

							  chngegetprice = Double.parseDouble(vr.getprice); 
							  
							  if(vr.scrip.equals(vr.getscrip))
							  {	
								  if(vr.action.equals("BUY"))							  
								  {
									  robot.keyPress(KeyEvent.VK_F1);
									  robot.keyRelease(KeyEvent.VK_F1);
								  }
								  if(vr.action.equals("SELL"))							  
								  {
									  robot.keyPress(KeyEvent.VK_F2);
									  robot.keyRelease(KeyEvent.VK_F2);
								  }							  
								  
								  w.findElement(By.id("1018")).sendKeys(vr.clientcode);
								  
								  robot.keyPress(KeyEvent.VK_TAB);
								  robot.keyPress(KeyEvent.VK_DELETE);	
								  
								  w.findElement(By.id("1400")).sendKeys(String.valueOf(vr.qty));
								  
								  ///capture actual price
								  robot.keyPress(KeyEvent.VK_TAB);
								  vr.getamt = w.findElement(By.id("1404")).getText();
								  System.out.println("getamount "+vr.getamt);							  
								
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
								  								  								  
								  robot.keyPress(KeyEvent.VK_DELETE);
								  
								  if((vr.roundoff<=0)||vr.perc==0)
								  {
									  vr.roundoff=0.0;									  
								  }									  
								  w.findElement(By.id("1404")).sendKeys(String.valueOf(vr.roundoff));
								  
								  w.findElement(By.id("1342")).click();/////new buy order button
								  Thread.sleep(1500);
								  robot.keyPress(KeyEvent.VK_ENTER);
								  Thread.sleep(1500);
								  w.findElement(By.id("2")).click();
								  
								  //{		
										  log = report1.startTest("Place Normal " +vr.scrip+ " Order");
										  log.log(LogStatus.INFO, "Scrip " +vr.scrip+ " is selected");
										  log.log(LogStatus.INFO, "Entered Client Code " +vr.clientcode);			
										  log.log(LogStatus.INFO, "Selected action is "+vr.action);			
										  log.log(LogStatus.INFO, "Entered quantity "+vr.qty);	
										  log.log(LogStatus.INFO, "Actual price is " +vr.getamt);
										  log.log(LogStatus.INFO, "Modified price is " +vr.roundoff);	
										  log.log(LogStatus.INFO, "Select Order Mark : " +ordermrk);
	  												  
///////////////////////////////////////////////////////
										  String msg2 = "Success. ABWS Order number is:";
										  for(int m=0;m<2;)
										  {		
											  WebElement msg11 = w.findElementByClassName("SysListView32");		
											  List <WebElement> colcount1 = msg11.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.DataItem')]"));
											  System.out.println("item count "+colcount1.size());
											  
											  for(z=1;z<=colcount1.size();z++)
											  {		  													  
													  List<WebElement> radio1 = w.findElements(By.xpath("//*[contains(@Column, '"+z+"')]"));
													  System.out.println("radio1 size "+radio1.size());
													  
														for(int i=1;i<radio1.size();i++)
														{	
															getmsg = radio1.get(i).getAttribute("Name");
															System.out.println("in loop gm...."+getmsg);
															
															  if(getmsg.startsWith(msg2))
															  {																  
																  remdot = getmsg.split("\\:",0);
																  getordno = remdot[1].split("\\.", 0);
																  System.out.println("get on "+getordno[0]);											  						  							  
																  
																  log.log(LogStatus.INFO, ""+getmsg);
																  log.log(LogStatus.INFO, "Order no " +getordno[0]+ " placed successfully");
																  System.out.println("Completed successfully");	
																  
																  System.out.println("passed loop ");
																  z=100;
																  break;																  
															  }
															  else
															  {
																  System.out.println("else loop ");//No success message " +getmsg);
															  }															  
														}																										
										  }
											  break;
										  }
										  break;
							  }
						  }
						  break;
					  }
				  }
				  else
				  {
					  System.out.println("flag is N");
				  }
			  }
		}
										  
		
		 @Test
		 public void placemodcanorder() throws IOException, Exception
		 {	 	
			  ///report creation
			  report1 = new ExtentReports("D:\\Windowapp\\mock.html");
			  log = report1.startTest("Login Details");
			  
			  ////call login function
			  Login();
			  
			  robot = new Robot();				  		  
			  
			  openWL();

			  //deltakesnapshot();	
			  
			  placeorder();
			  
			  ///open all reports			  
			  log = report1.startTest("Reports Screenshot ");
			  
		 	}
		 
		 
		 @Test(priority=4001)
			public void nouse()
			  {
				  report1.endTest(log);
				  report1.flush();
				  //w.get("D:\\Windowapp\\mock.html");				  
				  //w.get(srcfile);
			  }
		 
		}




/*
												 String msg2 = "Success. ABWS Order number is:";
												 for(int m=0;m<2;)
												 {					  				  
												  WebElement msg = w.findElementByClassName("SysListView32");
												  List <WebElement> colcount1 = msg.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.DataItem')]"));
												  System.out.println("item count "+colcount1.size());
												  
												  WebElement colcount2= msg.findElement(By.xpath("./*[contains(@ControlType, 'ControlType.DataItem')]"));
												  List <WebElement> txt = colcount2.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Text')]"));
												  System.out.println("txt count "+txt.size());
												  for(int y=1;y<=colcount1.size();y++)
												  {
													  System.out.println("y value "+y);
													  
													  String get = txt.get(y).getAttribute("Name");
													  System.out.println("get "+get);
													  for(int z=1;z<txt.size();)
													  {														  
														  getmsg = txt.get(z).getAttribute("Name");//.findElement(By.xpath("//*[contains(@Name), '"+msg2+"']")).getAttribute("Name");///
														  System.out.println("get txt message z "+getmsg);
															
														  if(getmsg.startsWith(msg2))
														  {
															  remdot = getmsg.split("\\:",0);
															  getordno = remdot[1].split("\\.", 0);
															  System.out.println("get on "+getordno[0]);											  						  							  
															  
															  log.log(LogStatus.INFO, ""+getmsg);
															  log.log(LogStatus.INFO, "Order no " +getordno[0]+ " placed successfully");
															  System.out.println("Completed successfully");	
															  break;
														  }
														  else
														  {
															  System.out.println("No success message "+getmsg);
															  break;
														  }
													  }
												  }
												  break;
												  //getmsg = msg.findElement(By.xpath("//*[contains(@Name), '"+msg2+"']")).getAttribute("Name");///working										  
											  }

*/



