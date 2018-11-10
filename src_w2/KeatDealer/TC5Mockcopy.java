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
import org.testng.TestListenerAdapter;
import org.testng.TestNG;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.LogStatus;

import Libraries.functions;
import Libraries.variables;
import Main.beforesuitedealer;

public class TC5Mockcopy extends beforesuitedealer
	{	
		public static String[] getordno,remdot,gettsloordno,tsloremdot,getsmordno,smremdot;
		public static double cnvrtcaputil,placecnvrtcaputil,matchround,chngegetprice,chngestrnggetamt;
		public static String caputilscrip,caputil,getmsg,capordno,getxchnge,TSLOchk,tsloordno,smordno;
		public static String ME,capME,rghclient,capscrip,delmark,getstat,nettrdval,getqty;
		WebElement keatpro,op;		
		public static int i,j,xcelrc,m,addqty,bpprice;
		Robot robot;
		public static float SLOprice,spread;
		public static String gettrigpr;
		public static double totbp;
		public static String bpr,sqty,bqty,spr,bord,sord;

		variables vr = new variables();
		functions fn = new functions();	
		
		public static void main(String[] args) 
		  {
			  TestListenerAdapter tla = new TestListenerAdapter();
			    TestNG testng = new TestNG();
			     Class[] classes = new Class[]
			     {
			    	 TC5Mockcopy.class     
			     };
			     testng.setTestClasses(classes);
			     testng.addListener(tla);
			     testng.run();		  	  
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
		}
		
		public void md() throws Exception
		{
			/*
			tslorpt();
			  
		  	WebElement nptable = w.findElement(By.name("TSLO Order Report"));//.findElement(By.id("37033"));			  
			List < WebElement > rc = nptable.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			System.out.println("tslo row count "+ rc.size());
			  
			 WebElement abc = nptable.findElement(By.name("Row " + (1) + ""));
			 capscrip = abc.findElement(By.name("Row " + (1) + ", Column 2")).getText();
			 
			for(int p=1;p<rc.size();p++)
			{		  				 
				 //capture ME for BSE
				 capME = abc.findElement(By.name("Row " + (p) + ", Column 13")).getText();
				 System.out.println("status tslo "+capME);
			}
			
			close();
			*/			
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
						  j=j+1;
						  ME = vr.s.getRow(i).getCell(j).getStringCellValue();
						  
						  for(int k=1;k<=rc1.size();k++)
						  {
							  vr.getscrip = wltable.findElement(By.name("Row " + (k) + ", Column 1")).getText();

							  if(vr.scrip.equals(vr.getscrip) && ME.equals(getxchnge))
							  {
								  wltable.findElement(By.name("Row " + (k)+ ", Column 1")).click();
								  
								  Thread.sleep(500);
								  robot.keyPress(KeyEvent.VK_F5);
								  robot.keyRelease(KeyEvent.VK_F5);
								  Thread.sleep(500);
								  
								  WebElement msg = w.findElementByClassName("SysListView32");
								  List <WebElement> colcount1 = msg.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.DataItem')]"));
								  System.out.println("item count "+colcount1.size());
								  
								  WebElement colcount2= msg.findElement(By.xpath("./*[contains(@ControlType, 'ControlType.DataItem')]"));
								  List <WebElement> txt = colcount2.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Text')]"));
								  System.out.println("txt count "+txt.size());
								  
								  for(int y=1;y<=colcount1.size();y++)
								  {
									  System.out.println("y value "+y);				  
									  for(int z=1;z<txt.size();)
									  {														  
										  bpr = txt.get(z).getAttribute("Name");//.findElement(By.xpath("//*[contains(@Name), '"+msg2+"']")).getAttribute("Name");///
										  System.out.println("b pr "+bpr);
										  z=z+1;
										  bqty = txt.get(z).getAttribute("Name");
										  System.out.println("b qty "+bqty);
										  z=z+1;
										  bord = txt.get(z).getAttribute("Name");
										  System.out.println("b ord "+bord);
										  z=z+1;
										  spr = txt.get(z).getAttribute("Name");
										  System.out.println("s pr "+spr);
										  z=z+1;
										  sqty = txt.get(z).getAttribute("Name");
										  System.out.println("s qty "+sqty);
										  z=z+1;
										  sord = txt.get(z).getAttribute("Name");
										  System.out.println("s ord "+sord);
										  break;
									  }
									  if(bpr !=null &&  !bpr.isEmpty())
									  {
										  robot.keyPress(KeyEvent.VK_F2);
										  robot.keyRelease(KeyEvent.VK_F2);
										  w.findElement(By.id("1018")).sendKeys("1Y308");
										  w.findElement(By.id("1400")).sendKeys(bqty);
										  w.findElement(By.id("1404")).sendKeys("0");
										  
										  w.findElement(By.id("1342")).click();/////new buy order button
										  Thread.sleep(1500);
										  
										  w.findElement(By.id("1")).click();/////success button
										  Thread.sleep(1500);							  
										  
										  //cancel								  
										  w.findElement(By.id("2")).click();
									  }
									  else
									  {
										  /*
										  robot.keyPress(KeyEvent.VK_F1);
										  robot.keyRelease(KeyEvent.VK_F1);
										  w.findElement(By.id("1018")).sendKeys("1Y305");
										  w.findElement(By.id("1400")).sendKeys(sqty);
										  w.findElement(By.id("1404")).sendKeys("0");
										  
										  w.findElement(By.id("1342")).click();/////new buy order button
										  Thread.sleep(1500);
										  
										  w.findElement(By.id("1")).click();/////success button
										  Thread.sleep(1500);							  
										  
										  //cancel								  
										  w.findElement(By.id("2")).click();
										  */
									  }									  
										  if(spr !=null &&  !spr.isEmpty())
										  {
											  robot.keyPress(KeyEvent.VK_F1);
											  robot.keyRelease(KeyEvent.VK_F1);
											  w.findElement(By.id("1018")).sendKeys("1Y306");
											  w.findElement(By.id("1400")).sendKeys(sqty);
											  w.findElement(By.id("1404")).sendKeys("0");
											  
											  w.findElement(By.id("1342")).click();/////new buy order button
											  Thread.sleep(1500);
											  
											  w.findElement(By.id("1")).click();/////success button
											  Thread.sleep(1500);							  
											  
											  //cancel								  
											  w.findElement(By.id("2")).click();
										  }
										  else
										  {
											  /*
											  robot.keyPress(KeyEvent.VK_F2);
											  robot.keyRelease(KeyEvent.VK_F2);
											  w.findElement(By.id("1018")).sendKeys("1Y302");
											  w.findElement(By.id("1400")).sendKeys(bqty);
											  w.findElement(By.id("1404")).sendKeys("0");
											  
											  w.findElement(By.id("1342")).click();/////new buy order button
											  Thread.sleep(1500);
											  
											  w.findElement(By.id("1")).click();/////success button
											  Thread.sleep(1500);							  
											  
											  //cancel								  
											  w.findElement(By.id("2")).click();
											  */
										  }
								  }
								  break;
							  }
							  else
							  {
								  System.out.println("scrip not found ");
							  }
							  Thread.sleep(1000);
							  robot.keyPress(KeyEvent.VK_ESCAPE);
							  Thread.sleep(1000);								 
							  robot.keyRelease(KeyEvent.VK_ESCAPE);
							  Thread.sleep(1000);
						  }
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
			  
			  //w.findElement(By.id("#32770")).click();
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_F4);
			  Thread.sleep(500);
			  robot.keyRelease(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  //robot.keyRelease(KeyEvent.VK_F4);
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
		
		public void smstatus() throws Exception
		{
			orderrpt();
			
			//count no of rows
			WebElement rpttable = w.findElement(By.name("Order Report [All]")).findElement(By.id("1003"));
			List < WebElement > rc3 = rpttable.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			System.out.println("row count "+ rc3.size());
											
			for(m=1;m<=rc3.size();m++)
			{
				System.out.println("get sm ord no "+getsmordno[0]);
				
				smordno = rpttable.findElement(By.name("Row " + (m) + ", Column 4")).getText();
				System.out.println("ord no "+smordno);				  
													  
				///modify order						  
				getstat = rpttable.findElement(By.name("Row " + (m)+ ", Column 11")).getText();
				System.out.println("sm status "+getstat);
													  
				if(smordno.equals(getsmordno[0]) && getstat.equals("TRAD"))
				{
					rpttable.findElement(By.name("Row " + (m)+ "")).click();
					log.log(LogStatus.INFO, "Select order no : "+smordno);
					log.log(LogStatus.INFO, "Super multiple plus Order No " +smordno+ " is Traded");
					break;
				}
				else //if(capordno.equals(getordno[0]) && getstat.equals("TRAD"))
					 {			
						log.log(LogStatus.INFO, "Order No " +smordno+ " is OPN");
					 }	
				}				
			  close();	
			
		}
		
		public void normalordstatus() throws Exception
		{
			orderrpt();
			
			//count no of rows
			WebElement rpttable = w.findElement(By.name("Order Report [All]")).findElement(By.id("1003"));
			List < WebElement > rc3 = rpttable.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			System.out.println("row count "+ rc3.size());
											
			for(m=1;m<=rc3.size();m++)
			{
				System.out.println("get ord no "+getordno[0]);
				
				capordno = rpttable.findElement(By.name("Row " + (m) + ", Column 4")).getText();
				System.out.println("ord no "+capordno);				  
													  
				///modify order						  
				getstat = rpttable.findElement(By.name("Row " + (m)+ ", Column 11")).getText();
				System.out.println("normal status "+getstat);
													  
				if(capordno.equals(getordno[0]) && getstat.equals("TRAD"))
				{
					rpttable.findElement(By.name("Row " + (m)+ "")).click();
					log.log(LogStatus.INFO, "Select order no : "+capordno);
					log.log(LogStatus.INFO, "Normal Order No " +capordno+ " is Traded");
					break;
				}
				else //if(capordno.equals(getordno[0]) && getstat.equals("TRAD"))
					 {			
						log.log(LogStatus.INFO, "Order No " +capordno+ " is OPN");
					 }	
				}				
			  close();	
			
		}
		
		public void tslostatus() throws Exception
		{
			tslorpt();
			
			System.out.println("get tslo ord no from status fn "+gettsloordno[0]);
			//count no of rows
			WebElement rpttable = w.findElement(By.name("TSLO Order Report")).findElement(By.id("1426"));
			List < WebElement > rc3 = rpttable.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			System.out.println("row count "+ rc3.size());
			int a =2;						
			for(m=2;m<=rc3.size();m++)
			{				
				System.out.println("in loop get tslo ord no from status fn "+gettsloordno[0]);
				
				tsloordno = rpttable.findElement(By.name("Row " + (a) + ", Column 1")).getText();
				System.out.println("tslo ord no "+tsloordno);				  
													  
				///modify order						  
				getstat = rpttable.findElement(By.name("Row " + (m)+ ", Column 13")).getText();
				System.out.println("tslo status "+getstat);
													  
				if(tsloordno.equals(gettsloordno[0]) && getstat.equals("TRAD"))
				{
					log.log(LogStatus.INFO, "TSLO Order No " +tsloordno+ " is TRAD ");					
				}
				else
				{	
					a=a+2;
					break;					 
				}	
			}				
			  close();	
			
		}
	  		
		public void smsquareoff() throws Exception
		{
			///for square off
			openposition();											  											  	
			
			//count no of rows
			WebElement nptable = w.findElement(By.name("Net Position")).findElement(By.id("37033"));			  
			List < WebElement > rc = nptable.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			System.out.println("netpos row count "+ rc.size());
			  
			for(int p=1;p<rc.size();p++)
			{				  
				 WebElement abc = nptable.findElement(By.name("Row " + (p) + ""));
				 capscrip = abc.findElement(By.name("Row " + (p) + ", Column 1")).getText();
				 //System.out.println("cap util "+capscrip);
				 
				 //capture ME for BSE
				 capME = abc.findElement(By.name("Row " + (p) + ", Column 2")).getText();
				 //System.out.println("cap ME "+capME);
				 
				 delmark = abc.findElement(By.name("Row " + (p) + ", Column 16")).getText();
				 //System.out.println("del mark "+delmark);
				 
				 nettrdval = abc.findElement(By.name("Row " + (p) + ", Column 9")).getText();
				 System.out.println("net trd val----- "+nettrdval);
				 
				 if(capscrip.equals(vr.scripchk) && capME.equals(getxchnge) && (delmark.equals("Super Multiple Plus")))
				 {			
					 if(nettrdval.equals("0.00"))
					 {
						 abc.findElement(By.name("Row " + (p) + ", Column 1")).click();
						 robot.mousePress(InputEvent.BUTTON3_DOWN_MASK);
						 Thread.sleep(500);
						 robot.mouseRelease(InputEvent.BUTTON3_DOWN_MASK);
						 Thread.sleep(500);
						 robot.keyPress(KeyEvent.VK_S);
						 Thread.sleep(500);
						 log.log(LogStatus.INFO, "Order is not traded so cannot be square off ");
						 break;
					 }						  						  
					  else
					  {
						  abc.findElement(By.name("Row " + (p) + ", Column 1")).click();
						  robot.mousePress(InputEvent.BUTTON3_DOWN_MASK);
						  Thread.sleep(500);
						  robot.mouseRelease(InputEvent.BUTTON3_DOWN_MASK);
						  Thread.sleep(500);
						  robot.keyPress(KeyEvent.VK_S);
						  Thread.sleep(500);
						  log.log(LogStatus.INFO, "Select super multiple plus order scrip: "+capscrip);	
		  
						  ////check for square off order button
						  w.findElement(By.id("1342")).click();/////square off button
						  Thread.sleep(1500);
						  //w.findElement(By.id("1")).click();/////success button
						  robot.keyPress(KeyEvent.VK_ENTER);
						  Thread.sleep(1500);
						  w.findElement(By.id("2")).click();
						  log.log(LogStatus.INFO, "Order is squared off ");
						  break;									  
					  }								  
				 }
				 else
				 {
					 //log.log(LogStatus.INFO, "No super multiple plus order available");								 
				 }
			}
			close();	
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
						  System.out.println("scrip "+vr.scrip);
						  vr.forreportscripname = vr.scrip;
						  						  
						  log.log(LogStatus.INFO, "Scrip " +vr.scrip+ " is selected");			
						  
						  ///excel for BSE
						  j=j+1;
						  ME = vr.s.getRow(i).getCell(j).getStringCellValue();
						  //System.out.println("ME "+ME);
						  		
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
						  String ordermrk = vr.s.getRow(i).getCell(j).getStringCellValue();

						  j=j+1;
						  TSLOchk = vr.s.getRow(i).getCell(j).getStringCellValue();
						  
						  j=j+1;
						  rghclient = vr.s.getRow(i).getCell(j).getStringCellValue();
						  
						  j=j+1;
						  addqty = (int) vr.s.getRow(i).getCell(j).getNumericCellValue();
						  
						  j=j+1;
						  SLOprice = (float) vr.s.getRow(i).getCell(j).getNumericCellValue();
						  
						  j=j+1;
						  vr.modqty = (int) vr.s.getRow(i).getCell(j).getNumericCellValue();
						  
						  j=j+1;
						  spread = (float) vr.s.getRow(i).getCell(j).getNumericCellValue();
						  
						  j=j+1;
						  bpprice = (int) vr.s.getRow(i).getCell(j).getNumericCellValue();
						  
						  if(TSLOchk.equals("ExistingTSLO"))
						  {
							  log = report1.startTest("Place Existing TSLO Order for :"+vr.forreportscripname);
							  ExistingTSLO();
							  break;
						  }
						  else if(TSLOchk.equals("ExistingTSLO-BP"))
						  {
							  log = report1.startTest("Place Existing Bracket TSLO Order for :"+vr.forreportscripname);
							  ExistingTSLOBP();
							  break;
						  }
						  
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
							  
							  if(vr.scrip.equals(vr.getscrip) && ME.equals(getxchnge))
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
								  //log.log(LogStatus.INFO, "Entered Client Code " +vr.clientcode);			
								  //log.log(LogStatus.INFO, "Selected action is "+vr.action);			
								  
								  robot.keyPress(KeyEvent.VK_TAB);
								  robot.keyPress(KeyEvent.VK_DELETE);	
								  
								  w.findElement(By.id("1400")).sendKeys(String.valueOf(vr.qty));
								  //log.log(LogStatus.INFO, "Entered quantity "+vr.qty);	
								  
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
								  //log.log(LogStatus.INFO, "Actual price is " +vr.getamt);
								  								  								  
								  robot.keyPress(KeyEvent.VK_DELETE);
								  
								  if((vr.roundoff<=0)||vr.perc==0)
								  {
									  vr.roundoff=0.0;									  
								  }									  
								  w.findElement(By.id("1404")).sendKeys(String.valueOf(vr.roundoff));
								  //log.log(LogStatus.INFO, "Modified price is " +vr.roundoff);	
						  
								  //log.log(LogStatus.INFO, "Select Order Mark : " +ordermrk);
								  
								  if(ordermrk.equals("Normal"))
								  {										  
									  if(TSLOchk.equals("WLTSLO"))
									  {
										  log = report1.startTest("Place TSLO Order for :"+vr.forreportscripname);
										  WLTSLO();
									  }
									  else if(TSLOchk.equals("WLTSLO-BP"))
									  {
										  log = report1.startTest("Place Bracket TSLO Order for :"+vr.forreportscripname);
										  WLTSLOBP();
									  }									  
									  else if(TSLOchk.equals("None"))
									  {
										  log = report1.startTest("Place Normal " +vr.scrip+ " Order");
										  log.log(LogStatus.INFO, "Entered Client Code " +vr.clientcode);			
										  log.log(LogStatus.INFO, "Selected action is "+vr.action);			
										  log.log(LogStatus.INFO, "Entered quantity "+vr.qty);	
										  log.log(LogStatus.INFO, "Actual price is " +vr.getamt);
										  log.log(LogStatus.INFO, "Modified price is " +vr.roundoff);	
										  log.log(LogStatus.INFO, "Select Order Mark : " +ordermrk);
										  
										  if(vr.roundoff!=0)
										  {
											  w.findElement(By.id("1342")).click();/////new buy order button
											  Thread.sleep(1500);
											  log.log(LogStatus.INFO, "Click on New "+vr.action+ " Order ");
											  
											  //w.findElement(By.id("1")).click();/////success button
											  robot.keyPress(KeyEvent.VK_ENTER);
											  Thread.sleep(1500);							  
											  
											  //cancel								  
											  w.findElement(By.id("2")).click();
											  log.log(LogStatus.INFO, "Click on Cancel ");
											  
											  String msg2 = "Success. ABWS Order number is:";
											  for(int m=0;m<2;)
											  {						  				  
												  WebElement msg = w.findElementByClassName("SysListView32");										  
												  getmsg = msg.findElement(By.xpath("//*[contains(@Name, '"+msg2+"')]")).getAttribute("Name");///working
												  System.out.println("get message "+getmsg);
												  if(getmsg.startsWith(msg2))
												  {
													  remdot = getmsg.split("\\:",0);
													  getordno = remdot[1].split("\\.", 0);
													  System.out.println("get on "+getordno[0]);											  						  							  
														  
													  log.log(LogStatus.INFO, ""+getmsg);
													  log.log(LogStatus.INFO, "Order no " +getordno[0]+ " placed successfully");
													  System.out.println("Completed successfully");	
													  
												  }
												  else
												  {
													  System.out.println("No success message ");
												  }
												  break;
											  }
											  Normalorder();
										  }
										  else if(vr.roundoff==0)
										  {
											  if(vr.perc !=0)
											  {												  
												  w.findElement(By.id("1342")).click();/////new buy order button
												  Thread.sleep(1500);
												  log.log(LogStatus.INFO, "Click on New "+vr.action+ " Order ");
												  
												  //w.findElement(By.id("1")).click();/////success button
												  robot.keyPress(KeyEvent.VK_ENTER);
												  Thread.sleep(1500);
												  
												  w.findElement(By.id("1018")).sendKeys(rghclient);
												  log.log(LogStatus.INFO, "Enter other Client " +rghclient);	
	
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
												  w.findElement(By.id("1400")).sendKeys(String.valueOf(vr.qty));
												  log.log(LogStatus.INFO, "Entered quantity " +vr.qty);	
	
												  w.findElement(By.id("1404")).sendKeys("0.00"); ///for trade at mkt
												  log.log(LogStatus.INFO, "Place order at market price ");	
	
												  w.findElement(By.id("1342")).click();/////new buy order button
												  Thread.sleep(1500);
												  //w.findElement(By.id("1")).click();/////success button
												  robot.keyPress(KeyEvent.VK_ENTER);
												  Thread.sleep(1500);
												  w.findElement(By.id("2")).click();
											  }
											  else
											  {
												  w.findElement(By.id("1342")).click();/////new buy order button
												  Thread.sleep(1500);
												  //w.findElement(By.id("1")).click();/////success button
												  robot.keyPress(KeyEvent.VK_ENTER);
												  Thread.sleep(1500);
												  w.findElement(By.id("2")).click();
												  log.log(LogStatus.INFO, "Place order at market price ");	

											  }
										  }
									  }	
								  }
								  else if(ordermrk.equals("Super Multiple Plus"))
								  {	
									  log = report1.startTest("Place Super Multiple Order for :"+vr.forreportscripname);
									  SuperMultiplePlus();
								  }								  
								  break;
							  }
							  else
							  {
								  System.out.println("scrip not found ");
							  }
						  }
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
		
		public void WLTSLO() throws Exception
		{
			String gettrigpr;
			//w.findElement(By.id("1404")).sendKeys("0.00");
			w.findElement(By.id("1421")).click();				
			w.findElement(By.id("1409")).sendKeys(String.valueOf(spread));
			log.log(LogStatus.INFO, "Spread is: " +spread);
			robot.keyPress(KeyEvent.VK_TAB);
			robot.keyRelease(KeyEvent.VK_TAB);
			gettrigpr = w.findElement(By.id("1411")).getText();
			System.out.println("tslo trig pr "+gettrigpr);
			w.findElement(By.id("1342")).click();/////new buy order button
			Thread.sleep(1500);
			//w.findElement(By.id("1")).click();/////success button
			  robot.keyPress(KeyEvent.VK_ENTER);
			Thread.sleep(1500);
			w.findElement(By.id("2")).click();
			log.log(LogStatus.INFO, "Place TSLO order from watchlist");	
			
			String msg2 = "Success.\n TSLO reference number is:";
			  for(int m=0;m<2;)
			  {						  				  
				  WebElement msg = w.findElementByClassName("SysListView32");										  
				  getmsg = msg.findElement(By.xpath("//*[contains(@Name, '"+msg2+"')]")).getAttribute("Name");///working
				  System.out.println("get message "+getmsg);
				  if(getmsg.startsWith(msg2))
				  {
					  tsloremdot = getmsg.split("\\:",0);
					  gettsloordno = tsloremdot[1].split("\\.", 0);
					  System.out.println("get on "+gettsloordno[0]);											  						  							  
						  
					  log.log(LogStatus.INFO, ""+getmsg);
					  log.log(LogStatus.INFO, "Order no " +gettsloordno[0]+ " placed successfully");
					  System.out.println("Completed successfully");	
				  }
				  else
				  {
					  System.out.println("No success message ");
				  }
				  break;
			  }
			
			//for leg1 trad	
			if(vr.bprice.equals("0") && vr.sprice.equals("0"))
			{
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
				  w.findElement(By.id("1018")).sendKeys(rghclient);
				  w.findElement(By.id("1400")).sendKeys(String.valueOf(vr.qty));
				  if(vr.action.equals("BUY"))
				  {
					  vr.matchround = Double.parseDouble(gettrigpr)-spread;
				  }
				  else
				  {
					  vr.matchround = Double.parseDouble(gettrigpr)+spread;
				  }
				  System.out.println("print calc ord pr "+vr.matchround);
				  w.findElement(By.id("1404")).sendKeys(gettrigpr);			  
				  w.findElement(By.id("1421")).click();			
				  w.findElement(By.id("1409")).sendKeys(String.valueOf(spread));
				  robot.keyPress(KeyEvent.VK_TAB);
				  robot.keyRelease(KeyEvent.VK_TAB);
				  gettrigpr = w.findElement(By.id("1411")).getText();
				  System.out.println("tslo trig pr "+gettrigpr);
				  log.log(LogStatus.INFO, "Spread is: " +spread);
				  w.findElement(By.id("1342")).click();/////new buy order button
				  Thread.sleep(1500);
				  //w.findElement(By.id("1")).click();/////success button
				  robot.keyPress(KeyEvent.VK_ENTER);
				  Thread.sleep(1500);
				  w.findElement(By.id("2")).click();	
				  log.log(LogStatus.INFO, "Place TSLO order from watchlist to trade leg1 ");	

			}
			  //for leg2 trade place normal order
			  if(vr.action.equals("SELL"))							  
			  {
				  robot.keyPress(KeyEvent.VK_CONTROL);
				  robot.keyPress(KeyEvent.VK_F2);
				  robot.keyRelease(KeyEvent.VK_CONTROL);
				  robot.keyRelease(KeyEvent.VK_F2);
			  }
			  else if(vr.action.equals("BUY"))							  
			  {
				  robot.keyPress(KeyEvent.VK_CONTROL);
				  robot.keyPress(KeyEvent.VK_F1);
				  robot.keyRelease(KeyEvent.VK_CONTROL);
				  robot.keyRelease(KeyEvent.VK_F1);
			  }
			  w.findElement(By.id("1018")).sendKeys("1Y302");
			  w.findElement(By.id("1400")).sendKeys(String.valueOf(vr.qty)); 
			  w.findElement(By.id("1404")).sendKeys(String.valueOf(gettrigpr));	  
			  w.findElement(By.id("1342")).click();/////new buy order button
			  Thread.sleep(1500);
			  //w.findElement(By.id("1")).click();/////success button
			  robot.keyPress(KeyEvent.VK_ENTER);
			  Thread.sleep(1500);
			  w.findElement(By.id("2")).click();
			  log.log(LogStatus.INFO, "Place normal order from watchlist to trade leg2 ");	

			  //tslorpt();
			  //Thread.sleep(5000);
			  //close();
		}
		
		public void WLTSLOBP() throws Exception
		{			
			//w.findElement(By.id("1404")).sendKeys("0.00");
			w.findElement(By.id("1421")).click();				
			w.findElement(By.id("1409")).sendKeys(String.valueOf(spread));
			log.log(LogStatus.INFO, "Spread is: " +spread);
			robot.keyPress(KeyEvent.VK_TAB);
			robot.keyRelease(KeyEvent.VK_TAB);
			gettrigpr = w.findElement(By.id("1411")).getText();
			System.out.println("tslo trig pr "+gettrigpr);
			w.findElement(By.id("1132")).click();	
			if(vr.action.equals("BUY"))
			{
				totbp = Double.parseDouble(gettrigpr) + bpprice;
			}
			else
			{
				totbp = Double.parseDouble(gettrigpr) - bpprice;
			}
			System.out.println("tot bp "+totbp);
			w.findElement(By.id("1418")).sendKeys(String.valueOf(totbp));
			w.findElement(By.id("1342")).click();/////new buy order button
			Thread.sleep(1500);
			//w.findElement(By.id("1")).click();/////success button
			  robot.keyPress(KeyEvent.VK_ENTER);
			Thread.sleep(1500);
			w.findElement(By.id("2")).click();
			log.log(LogStatus.INFO, "Place TSLO order from watchlist with book profit ");	

			String msg2 = "Success.\n TSLO reference number is:";
			  for(int m=0;m<2;)
			  {						  				  
				  WebElement msg = w.findElementByClassName("SysListView32");										  
				  getmsg = msg.findElement(By.xpath("//*[contains(@Name, '"+msg2+"')]")).getAttribute("Name");///working
				  System.out.println("get message "+getmsg);
				  if(getmsg.startsWith(msg2))
				  {
					  tsloremdot = getmsg.split("\\:",0);
					  gettsloordno = tsloremdot[1].split("\\.", 0);
					  System.out.println("get on "+gettsloordno[0]);											  						  							  
						  
					  log.log(LogStatus.INFO, ""+getmsg);
					  log.log(LogStatus.INFO, "Order no " +gettsloordno[0]+ " placed successfully");
					  System.out.println("Completed successfully");	
				  }
				  else
				  {
					  System.out.println("No success message ");
				  }
				  break;
			  }
			  
			//for leg1 trad
			if(vr.bprice.equals("0") && vr.sprice.equals("0"))
			{
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
				  w.findElement(By.id("1018")).sendKeys(rghclient);
				  w.findElement(By.id("1400")).sendKeys(String.valueOf(vr.qty));
				  if(vr.action.equals("BUY"))
				  {
					  vr.matchround = Double.parseDouble(gettrigpr)-spread;
				  }
				  else
				  {
					  vr.matchround = Double.parseDouble(gettrigpr)+spread;
				  }
				  System.out.println("print calc ord pr "+vr.matchround);
				  w.findElement(By.id("1404")).sendKeys(gettrigpr);			  
				  w.findElement(By.id("1421")).click();			
				  w.findElement(By.id("1409")).sendKeys(String.valueOf(spread));
				  robot.keyPress(KeyEvent.VK_TAB);
				  robot.keyRelease(KeyEvent.VK_TAB);
				  gettrigpr = w.findElement(By.id("1411")).getText();
				  System.out.println("tslo trig pr "+gettrigpr);
				  log.log(LogStatus.INFO, "Spread is: " +spread);
				  w.findElement(By.id("1132")).click();	
				  if(vr.action.equals("BUY"))
				  {
					totbp = Double.parseDouble(gettrigpr) + bpprice;
				  }
				  else
				  {
					totbp = Double.parseDouble(gettrigpr) - bpprice;
				  }
				  System.out.println("tot bp "+totbp);
				  w.findElement(By.id("1418")).sendKeys(String.valueOf(totbp));
				  w.findElement(By.id("1342")).click();/////new buy order button
				  Thread.sleep(1500);
				  //w.findElement(By.id("1")).click();/////success button
				  robot.keyPress(KeyEvent.VK_ENTER);
				  Thread.sleep(1500);
				  w.findElement(By.id("2")).click();		
				  log.log(LogStatus.INFO, "Place TSLO order from watchlist with book profit to trade leg1 ");	

			}
			  //for leg2 trade place normal order
			  if(vr.action.equals("SELL"))							  
			  {
				  robot.keyPress(KeyEvent.VK_CONTROL);
				  robot.keyPress(KeyEvent.VK_F2);
				  robot.keyRelease(KeyEvent.VK_CONTROL);
				  robot.keyRelease(KeyEvent.VK_F2);
			  }
			  else if(vr.action.equals("BUY"))							  
			  {
				  robot.keyPress(KeyEvent.VK_CONTROL);
				  robot.keyPress(KeyEvent.VK_F1);
				  robot.keyRelease(KeyEvent.VK_CONTROL);
				  robot.keyRelease(KeyEvent.VK_F1);
			  }
			  w.findElement(By.id("1018")).sendKeys("1Y302");
			  w.findElement(By.id("1400")).sendKeys(String.valueOf(vr.qty)); 
			  w.findElement(By.id("1404")).sendKeys(String.valueOf(totbp));	  
			  w.findElement(By.id("1342")).click();/////new buy order button
			  Thread.sleep(1500);
			  //w.findElement(By.id("1")).click();/////success button
			  robot.keyPress(KeyEvent.VK_ENTER);
			  Thread.sleep(1500);
			  w.findElement(By.id("2")).click();
			  log.log(LogStatus.INFO, "Place normal order from watchlist with book profit to trade leg2 ");	

			  //tslorpt();
			  //Thread.sleep(5000);
			  //close();			  
		}
		
		public void ExistingTSLO() throws Exception
		{
			openposition();
			
			//count no of rows
			WebElement nptable = w.findElement(By.name("Net Position")).findElement(By.id("37033"));			  
			List < WebElement > rc = nptable.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			System.out.println("netpos row count "+ rc.size());
			  
			for(int p=1;p<rc.size();p++)
			{				  
				 WebElement abc = nptable.findElement(By.name("Row " + (p) + ""));
				 capscrip = abc.findElement(By.name("Row " + (p) + ", Column 1")).getText();
				 //System.out.println("cap util "+capscrip);
				 remdot = capscrip.split("\\~",0);
				 System.out.println("cap util "+remdot[0]);
				 
				 //capture ME for BSE
				 capME = abc.findElement(By.name("Row " + (p) + ", Column 2")).getText();
				 //System.out.println("cap ME "+capME);
				 
				 delmark = abc.findElement(By.name("Row " + (p) + ", Column 16")).getText();
				 //System.out.println("del mark "+delmark);
				 
				 nettrdval = abc.findElement(By.name("Row " + (p) + ", Column 9")).getText();
				 System.out.println("net trd val----- "+nettrdval);
				 
				 //if(capscrip.equals(vr.getscrip) && capME.equals(getxchnge) && (delmark.equals("Normal")))
				 if(delmark.equals("Normal"))
				 {			
					 if(nettrdval.equals("0.00"))
					 {
						 abc.findElement(By.name("Row " + (p) + ", Column 1")).click();
						 robot.mousePress(InputEvent.BUTTON3_DOWN_MASK);
						 Thread.sleep(500);
						 robot.mouseRelease(InputEvent.BUTTON3_DOWN_MASK);
						 Thread.sleep(500);
						 robot.keyPress(KeyEvent.VK_S);
						 Thread.sleep(500);
						 log.log(LogStatus.INFO, "Order is not traded so cannot place Existing TSLO order ");
						 break;
					 }						  						  
					  else
					  {
						  abc.findElement(By.name("Row " + (p) + ", Column 1")).click();
						  robot.mousePress(InputEvent.BUTTON3_DOWN_MASK);
						  Thread.sleep(500);
						  robot.mouseRelease(InputEvent.BUTTON3_DOWN_MASK);
						  Thread.sleep(500);
						  robot.keyPress(KeyEvent.VK_S);
						  Thread.sleep(500);
						  log.log(LogStatus.INFO, "Select scrip: "+capscrip);	

						  ////place existing tslo order
						  getqty = w.findElement(By.id("1400")).getText();
						  System.out.println("get qty "+getqty);

						  w.findElement(By.id("1421")).click();				
						  w.findElement(By.id("1409")).sendKeys(String.valueOf(spread));
						  log.log(LogStatus.INFO, "Spread is: " +spread);
						  robot.keyPress(KeyEvent.VK_TAB);
						  robot.keyRelease(KeyEvent.VK_TAB);
						  gettrigpr = w.findElement(By.id("1411")).getText();
						  System.out.println("tslo trig pr "+gettrigpr);
						  w.findElement(By.id("1342")).click();/////square off button
						  Thread.sleep(1500);
						  //w.findElement(By.id("1")).click();/////success button
						  robot.keyPress(KeyEvent.VK_ENTER);
						  Thread.sleep(1500);
						  w.findElement(By.id("2")).click();
						  log.log(LogStatus.INFO, "Place Existing TSLO order from open position ");	
						  
						  String msg2 = "Success.\n TSLO reference number is:";
						  for(int m=0;m<2;)
						  {						  				  
							  WebElement msg = w.findElementByClassName("SysListView32");										  
							  getmsg = msg.findElement(By.xpath("//*[contains(@Name, '"+msg2+"')]")).getAttribute("Name");///working
							  System.out.println("get message "+getmsg);
							  if(getmsg.startsWith(msg2))
							  {
								  tsloremdot = getmsg.split("\\:",0);
								  gettsloordno = tsloremdot[1].split("\\.", 0);
								  System.out.println("get on "+gettsloordno[0]);											  						  							  
									  
								  log.log(LogStatus.INFO, ""+getmsg);
								  log.log(LogStatus.INFO, "Order no " +gettsloordno[0]+ " placed successfully");
								  System.out.println("Completed successfully");	
							  }
							  else
							  {
								  System.out.println("No success message ");
							  }
							  break;
						  }
						  
						  //for leg2 trade place normal order
						  if(Double.parseDouble(nettrdval)>0)
						  {
							  robot.keyPress(KeyEvent.VK_F1);
							  robot.keyRelease(KeyEvent.VK_F1);
						  }
						  else if(Double.parseDouble(nettrdval)<0)
						  {
							  robot.keyPress(KeyEvent.VK_F2);
							  robot.keyRelease(KeyEvent.VK_F2);
						  }
						  //for leg2 trade place normal order
						  w.findElement(By.id("1018")).sendKeys(rghclient);
						  w.findElement(By.id("1401")).sendKeys(remdot[0]);
						  w.findElement(By.id("1399")).sendKeys(capME);
						  robot.keyPress(KeyEvent.VK_TAB);
						  robot.keyPress(KeyEvent.VK_TAB);
						  robot.keyPress(KeyEvent.VK_DELETE);
						  robot.keyRelease(KeyEvent.VK_TAB);
						  robot.keyRelease(KeyEvent.VK_DELETE);
						  w.findElement(By.id("1400")).sendKeys(getqty); 
						  w.findElement(By.id("1404")).sendKeys(String.valueOf(gettrigpr));	  
						  w.findElement(By.id("1342")).click();/////new buy order button
						  Thread.sleep(1500);
						  //w.findElement(By.id("1")).click();/////success button
						  robot.keyPress(KeyEvent.VK_ENTER);
						  Thread.sleep(1500);
						  w.findElement(By.id("2")).click();	
						  log.log(LogStatus.INFO, "Place normal order from open position to trade leg2");	

						  //tslorpt();
						  //Thread.sleep(5000);
						  //close();
						  break;									  
					  }								  
				 }
				 else
				 {
					 //log.log(LogStatus.INFO, "No super multiple plus order available");								 
				 }
			}
			close();
		}
		
		public void ExistingTSLOBP() throws Exception
		{
			openposition();
			
			//count no of rows
			WebElement nptable = w.findElement(By.name("Net Position")).findElement(By.id("37033"));			  
			List < WebElement > rc = nptable.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			System.out.println("netpos row count "+ rc.size());
			  
			for(int p=1;p<rc.size();p++)
			{				  
				 WebElement abc = nptable.findElement(By.name("Row " + (p) + ""));
				 capscrip = abc.findElement(By.name("Row " + (p) + ", Column 1")).getText();
				 //System.out.println("cap util "+capscrip);
				 remdot = capscrip.split("\\~",0);
				 System.out.println("cap util "+remdot[0]);

				 //capture ME for BSE
				 capME = abc.findElement(By.name("Row " + (p) + ", Column 2")).getText();
				 //System.out.println("cap ME "+capME);
				 
				 delmark = abc.findElement(By.name("Row " + (p) + ", Column 16")).getText();
				 //System.out.println("del mark "+delmark);
				 
				 nettrdval = abc.findElement(By.name("Row " + (p) + ", Column 9")).getText();
				 System.out.println("net trd val----- "+nettrdval);
				 
				 if(delmark.equals("Normal"))
				 {			
					 if(nettrdval.equals("0.00"))
					 {
						 abc.findElement(By.name("Row " + (p) + ", Column 1")).click();
						 robot.mousePress(InputEvent.BUTTON3_DOWN_MASK);
						 Thread.sleep(500);
						 robot.mouseRelease(InputEvent.BUTTON3_DOWN_MASK);
						 Thread.sleep(500);
						 robot.keyPress(KeyEvent.VK_S);
						 Thread.sleep(500);
						 log.log(LogStatus.INFO, "Order is not traded so cannot place Existing TSLO order ");
						 break;
					 }						  						  
					  else
					  {
						  abc.findElement(By.name("Row " + (p) + ", Column 1")).click();
						  robot.mousePress(InputEvent.BUTTON3_DOWN_MASK);
						  Thread.sleep(500);
						  robot.mouseRelease(InputEvent.BUTTON3_DOWN_MASK);
						  Thread.sleep(500);
						  robot.keyPress(KeyEvent.VK_S);
						  Thread.sleep(500);
						  log.log(LogStatus.INFO, "Select scrip: "+capscrip);	
  
						  getqty = w.findElement(By.id("1400")).getText();
						  System.out.println("get qty "+getqty);

						  ////place existing tslo order
						  w.findElement(By.id("1421")).click();				
						  w.findElement(By.id("1409")).sendKeys(String.valueOf(spread));
						  log.log(LogStatus.INFO, "Spread is: " +spread);
						  robot.keyPress(KeyEvent.VK_TAB);
						  robot.keyRelease(KeyEvent.VK_TAB);
						  gettrigpr = w.findElement(By.id("1411")).getText();
						  System.out.println("tslo trig pr "+gettrigpr);
						  w.findElement(By.id("1132")).click();	
						  if(Double.parseDouble(nettrdval)>0)
						  {
							totbp = Double.parseDouble(gettrigpr) + bpprice;
						  }
						  else
						  {
							totbp = Double.parseDouble(gettrigpr) - bpprice;
						  }
						  System.out.println("tot bp "+totbp);
						  w.findElement(By.id("1418")).sendKeys(String.valueOf(totbp));
						  w.findElement(By.id("1342")).click();/////square off button
						  Thread.sleep(1500);
						  //w.findElement(By.id("1")).click();/////success button
						  robot.keyPress(KeyEvent.VK_ENTER);
						  Thread.sleep(1500);
						  w.findElement(By.id("2")).click();
						  log.log(LogStatus.INFO, "Place Existing bracket TSLO order from open position ");	
						  
						  String msg2 = "Success.\n TSLO reference number is:";
						  for(int m=0;m<2;)
						  {						  				  
							  WebElement msg = w.findElementByClassName("SysListView32");										  
							  getmsg = msg.findElement(By.xpath("//*[contains(@Name, '"+msg2+"')]")).getAttribute("Name");///working
							  System.out.println("get message "+getmsg);
							  if(getmsg.startsWith(msg2))
							  {
								  tsloremdot = getmsg.split("\\:",0);
								  gettsloordno = tsloremdot[1].split("\\.", 0);
								  System.out.println("get on "+gettsloordno[0]);											  						  							  
									  
								  log.log(LogStatus.INFO, ""+getmsg);
								  log.log(LogStatus.INFO, "Order no " +gettsloordno[0]+ " placed successfully");
								  System.out.println("Completed successfully");	
							  }
							  else
							  {
								  System.out.println("No success message ");
							  }
							  break;
						  }
						  
						  //for leg2 trade place normal order
						  if(Double.parseDouble(nettrdval)>0)
						  {
							  robot.keyPress(KeyEvent.VK_F1);
							  robot.keyRelease(KeyEvent.VK_F1);
						  }
						  else if(Double.parseDouble(nettrdval)<0)
						  {
							  robot.keyPress(KeyEvent.VK_F2);
							  robot.keyRelease(KeyEvent.VK_F2);
						  }
						  w.findElement(By.id("1018")).sendKeys(rghclient);
						  w.findElement(By.id("1401")).sendKeys(remdot[0]);
						  w.findElement(By.id("1399")).sendKeys(capME);
						  robot.keyPress(KeyEvent.VK_TAB);
						  robot.keyPress(KeyEvent.VK_TAB);
						  robot.keyPress(KeyEvent.VK_DELETE);
						  robot.keyRelease(KeyEvent.VK_TAB);
						  robot.keyRelease(KeyEvent.VK_DELETE);
						  w.findElement(By.id("1400")).sendKeys(getqty); 
						  w.findElement(By.id("1404")).sendKeys(String.valueOf(totbp));	  
						  w.findElement(By.id("1342")).click();/////new buy order button
						  Thread.sleep(1500);
						  //w.findElement(By.id("1")).click();/////success button
						  robot.keyPress(KeyEvent.VK_ENTER);
						  Thread.sleep(1500);
						  w.findElement(By.id("2")).click();
						  log.log(LogStatus.INFO, "Place normal order to trade leg2 ");	

						  //tslorpt();
						  //Thread.sleep(5000);
						  //close();
						  break;									  
					  }								  
				 }
				 else
				 {
					 //log.log(LogStatus.INFO, "No super multiple plus order available");								 
				 }
			}
			close();
		}
		
		public void Normalorder() throws Exception
		{
			  if(vr.roundoff!=0)
			  {	  				  
				  Normalmodify();
				  //capordno = getordno[0];
				  Normalcancel();
			  }
			  else if(vr.roundoff==0)
			  {
				  log.log(LogStatus.INFO, "Order traded");
			  }
		}
		
		public void Normalmodify() throws Exception
		{
			///////////////manage order
			orderrpt();
												
			//count no of rows
			WebElement rpttable = w.findElement(By.name("Order Report [All]")).findElement(By.id("1003"));
			List < WebElement > rc3 = rpttable.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			System.out.println("row count "+ rc3.size());
											
			for(m=1;m<=rc3.size();m++)
			{
				capordno = rpttable.findElement(By.name("Row " + (m) + ", Column 4")).getText();
				System.out.println("ord no "+capordno);				  
													  
				///modify order						  
				getstat = rpttable.findElement(By.name("Row " + (m)+ ", Column 11")).getText();
				System.out.println("mod status "+getstat);
													  
				if(capordno.equals(getordno[0]) && getstat.equals("OPN") || getstat.equals("AMO"))
				{
					rpttable.findElement(By.name("Row " + (m)+ "")).click();
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
														  
					w.findElement(By.id("1400")).sendKeys(String.valueOf(vr.modqty));
					log.log(LogStatus.INFO, "Modified quantity : "+vr.modqty);
														  
					w.findElement(By.id("1342")).click();
					Thread.sleep(1500);			  			  
					//w.findElement(By.id("1")).click();
					  robot.keyPress(KeyEvent.VK_ENTER);
					Thread.sleep(1500);
					w.findElement(By.id("2")).click();	///cancel modify dialog box		
					log.log(LogStatus.INFO, "Order no " +getordno[0]+ " modified successfully");
					log.log(LogStatus.INFO, "Modified Required Margin is : "+vr.roundoff1);	
					break;
				}
				else //if(capordno.equals(getordno[0]) && getstat.equals("TRAD"))
					 {			
						log.log(LogStatus.INFO, "Order No " +capordno+ " is not OPN so cannot modify");
					 }	
				}				
			  close();			
		}
		public void Normalcancel() throws Exception
		{
			  /////////cancel	
			  orderrpt();

			  WebElement rpttablecan = w.findElement(By.name("Order Report [All]")).findElement(By.id("1003"));
			  List < WebElement > rc4 = rpttablecan.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			  System.out.println("row count "+ rc4.size());

			  for(int c=1;c<=rc4.size();c++)
			  {
				  	System.out.println("get ord no "+getordno[0]);
					capordno = rpttablecan.findElement(By.name("Row " + (c) + ", Column 4")).getText();
					System.out.println("ord no "+capordno);				  
														  
					///modify order						  
					getstat = rpttablecan.findElement(By.name("Row " + (c)+ ", Column 11")).getText();
					System.out.println("mod status "+getstat);

				  if(capordno.equals(getordno[0]) && getstat.equals("OPN") || getstat.equals("AMO"))
				  {
					  	rpttablecan.findElement(By.name("Row " + (c)+ "")).click();
						log.log(LogStatus.INFO, "Select order no : "+capordno);
						Thread.sleep(500);
						robot.mousePress(InputEvent.BUTTON3_DOWN_MASK);
						Thread.sleep(500);
						robot.mouseRelease(InputEvent.BUTTON3_DOWN_MASK);
						Thread.sleep(500);
						robot.keyPress(KeyEvent.VK_C);
						Thread.sleep(500);
						w.findElement(By.id("1342")).click();
						Thread.sleep(1500);			  			  
						//w.findElement(By.id("1")).click();
						robot.keyPress(KeyEvent.VK_ENTER);
						Thread.sleep(1500);
						w.findElement(By.id("2")).click();	///cancel cancel dialog box
						log.log(LogStatus.INFO, "Order no " +getordno[0]+ " cancelled successfully");	
						break;
				  }	
				  else //if(capordno.equals(getordno[0]) && getstat.equals("TRAD"))
				  {
					  log.log(LogStatus.INFO, "Order No " +capordno+ " is not OPN so cannot cancel");
				  }											
			  }
			  close();
		}
		
		public void SuperMultiplePlus() throws Exception
		{
			w.findElement(By.id("1014")).click();

			if(vr.roundoff!=0)
			  {
				  w.findElement(By.name("Super Multiple Plus")).click();
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
				  System.out.println("slo price " +vr.roundoff1);
				  log.log(LogStatus.INFO, "SLO Trigger Price is " +vr.roundoff1);
				  w.findElement(By.id("1342")).click();/////new buy order button
				  Thread.sleep(3000);
				  //w.findElement(By.id("1")).click();/////success button
				  robot.keyPress(KeyEvent.VK_ENTER);
				  Thread.sleep(3000);
				  w.findElement(By.id("2")).click();
				  log.log(LogStatus.INFO, "Place super multiple plus order ");	

				  String msg2 = "Success. ABWS Order number is:";
				  for(int m=0;m<2;)
				  {						  				  
					  WebElement msg = w.findElementByClassName("SysListView32");										  
					  getmsg = msg.findElement(By.xpath("//*[contains(@Name, '"+msg2+"')]")).getAttribute("Name");///working
					  System.out.println("get message "+getmsg);
					  if(getmsg.startsWith(msg2))
					  {
						  smremdot = getmsg.split("\\:",0);
						  getsmordno = smremdot[1].split("\\.", 0);
						  System.out.println("get on "+getsmordno[0]);											  						  							  
							  
						  log.log(LogStatus.INFO, ""+getmsg);
						  log.log(LogStatus.INFO, "Order no " +getsmordno[0]+ " placed successfully");
						  System.out.println("Completed successfully");	
					  }
					  else
					  {
						  System.out.println("No success message ");
					  }
					  break;
				  }
//////////////////////////////////				  
				  
					  //break;
			  }
			  else if(vr.roundoff==0)
			  {
				  for(int a=1;a<=1;) 
				  {
					  w.findElement(By.id("2")).click();
					  
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
					  w.findElement(By.id("1018")).sendKeys("1Y302");
					  w.findElement(By.id("1400")).sendKeys(String.valueOf(addqty));
					  w.findElement(By.id("1014")).click();
					  w.findElement(By.name("Normal")).click();							
					  w.findElement(By.id("1342")).click();/////new buy order button
					  Thread.sleep(3000);
					  //w.findElement(By.id("1")).click();/////success button
					  robot.keyPress(KeyEvent.VK_ENTER);
					  Thread.sleep(3000);
					  w.findElement(By.id("2")).click();
					  log.log(LogStatus.INFO, "Place order from rough client ");	
					  
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
					  w.findElement(By.id("1400")).sendKeys(String.valueOf(vr.qty));
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
					  Thread.sleep(3000);
					  //w.findElement(By.id("1")).click();/////success button
					  robot.keyPress(KeyEvent.VK_ENTER);
					  Thread.sleep(3000);
					  w.findElement(By.id("2")).click();	
					  log.log(LogStatus.INFO, "Place super multiple plus order ");
					  
					  String msg2 = "Success. ABWS Order number is:";
					  for(int m=0;m<2;)
					  {						  				  
						  WebElement msg = w.findElementByClassName("SysListView32");										  
						  getmsg = msg.findElement(By.xpath("//*[contains(@Name, '"+msg2+"')]")).getAttribute("Name");///working
						  System.out.println("get message "+getmsg);
						  if(getmsg.startsWith(msg2))
						  {
							  smremdot = getmsg.split("\\:",0);
							  getsmordno = smremdot[1].split("\\.", 0);
							  System.out.println("get on "+getsmordno[0]);											  						  							  
								  
							  log.log(LogStatus.INFO, ""+getmsg);
							  log.log(LogStatus.INFO, "Order no " +getsmordno[0]+ " placed successfully");
							  System.out.println("Completed successfully");	
						  }
						  else
						  {
							  System.out.println("No success message ");
						  }
						  break;
					  }
					  
					  break;
				}
			  }
			
////////////////////////////////////////
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
			  w.findElement(By.id("1018")).sendKeys(rghclient);
			  w.findElement(By.id("1400")).sendKeys(String.valueOf(vr.qty));
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
			  Thread.sleep(3000);
			  //w.findElement(By.id("1")).click();/////success button
			  robot.keyPress(KeyEvent.VK_ENTER);
			  Thread.sleep(3000);
			  w.findElement(By.id("2")).click();
			  log.log(LogStatus.INFO, "Trade super multiple plus order ");	
			  
////////////////////////////////////
			  ////assign scrip for smp
			  vr.scripchk = vr.getscrip;
		}
		
		 @Test
		 public void placemodcanorder() throws IOException, Exception
		 {	 	
			  ///report creation
			  report1 = new ExtentReports("D:\\testfile\\mock.html");
			  log = report1.startTest("Login Details");
			  
			  ////call login function
			  Login();
			  
			  robot = new Robot();	
			  
			  ////open watchlist
			  openWL();
			  
			  deltakesnapshot();
			  
			  //place order function
			  placeorder();
			  md();
			  smsquareoff();
			  //smstatus();
			  //normalordstatus();
			  
			  //tslostatus();
			  
			  ///open all reports			  
			  //vr.clientcode="1Y315";
			  //Thread.sleep(3000);
			  log = report1.startTest("Reports Screenshot ");
			  	 
			  openposition();
  			  Thread.sleep(2000);	
  			  passtakesnapshot();
			  close();
			  			  
			  todaysposition();
  			  Thread.sleep(2000);	
			  passtakesnapshot();
			  close();			  
			  
			  orderrpt();
  			  Thread.sleep(2000);	
			  passtakesnapshot();
			  close();
			  			  
			  traderpt();
  			  Thread.sleep(2000);	
			  passtakesnapshot();
  			  close();
			  
  			  limitspg();
  			  Thread.sleep(2000);	
			  passtakesnapshot();
   			  close();
   			  
   			  tslorpt();
			  Thread.sleep(10000);	
			  passtakesnapshot();
			  //close();
			  
			  copy2pdf();			  
			  
			  exit();
		 	}
		 
		 
		 @Test(priority=4001)
			public void nouse()
			  {
				  report1.endTest(log);
				  report1.flush();
				  //w.get("D:\\testfile\\mock.html");				  
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



