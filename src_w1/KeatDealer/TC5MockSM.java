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

public class TC5MockSM extends beforesuitedealer
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
		
		public void success() throws Exception
		{			  			  
			  w.findElement(By.id("1342")).click();/////new buy order button
			  Thread.sleep(2000);	  

			  //WebDriverWait wait= new WebDriverWait (w,30);
			  //WebElement element2=wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("OK")));
			  //element2.click();
			  robot.keyPress(KeyEvent.VK_ENTER);
			  Thread.sleep(2000);
			  
			  //w.findElement(By.id("1")).click();/////success button
			  //Thread.sleep(2000);	  
			  
			  //cancel								  
			  w.findElement(By.id("2")).click();
			  Thread.sleep(2000);	  
		}
		
		public void md() throws Exception
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
										  
										  success();										  
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
											  
											  success();											  
										  }
										  else
										  {
											  /*
											  robot.keyPress(KeyEvent.VK_F2);
											  robot.keyRelease(KeyEvent.VK_F2);
											  w.findElement(By.id("1018")).sendKeys(roughclient);
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
							  Thread.sleep(500);								 
							  robot.keyRelease(KeyEvent.VK_ESCAPE);
							  Thread.sleep(500);
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
			  
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_ALT);
			  Thread.sleep(500);
			  robot.keyPress(KeyEvent.VK_F4);
			  Thread.sleep(500);
			  robot.keyRelease(KeyEvent.VK_ALT);
			  Thread.sleep(500);
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
			  
			  limitext();			  
		}
		
		public void limitext() throws Exception
		{			
			WebElement limittable = w.findElement(By.name("Net Position")).findElement(By.id("1011"));
			  WebElement colcountlt = limittable.findElement(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));		
			  List < WebElement > rcnp = limittable.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			  System.out.println("row count "+ rcnp.size());
			  	
			  for(int x=1;x<rcnp.size();x++)
			  {
				  String gettotmargin = limittable.findElement(By.name("Row " + (x) + ", Column 1")).getText();
				  System.out.println("get margin text "+gettotmargin);											  
				  if(gettotmargin.equals("Total Margin"))
				  {
					  limittable.findElement(By.name("Row " + (x) + "")).click();												  
				  }
				  
				  if(gettotmargin.equals("Total Margin (Margin Trading)"))
				  {
					  limittable.findElement(By.name("Row " + (x) + "")).click();		
					  break;
				  }
			  }
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
		
		public void normmsg() throws Exception
		{
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
						  
							for(int i=0;i<radio1.size();i++)
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
									  storenormordno = getordno[0];
									  z=100;
									  break;																  
								  }
								  else
								  {
									  System.out.println("else loop ");//No success message " +getmsg);
									  log.log(LogStatus.INFO, ""+radio1.get(0).getAttribute("Name"));
								  }
								  
							}																				
			  }
				  break;
			  }
		}
		
		public void smmsg() throws Exception
		{
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
						  
							for(int i=0;i<radio1.size();i++)
							{	
								getmsg = radio1.get(i).getAttribute("Name");
								System.out.println("in loop gm...."+getmsg);
								
								  if(getmsg.startsWith(msg2))
								  {																  
									  remdot = getmsg.split("\\:",0);
									  getsmordno = remdot[1].split("\\.", 0);
									  System.out.println("get on "+getsmordno[0]);											  						  							  
									  
									  log.log(LogStatus.INFO, ""+getmsg);
									  log.log(LogStatus.INFO, "Order no " +getsmordno[0]+ " placed successfully");
									  System.out.println("Completed successfully");	
									  
									  System.out.println("passed loop ");
									  storesmordno = getsmordno[0];
									  z=100;
									  break;																  
								  }
								  else
								  {
									  System.out.println("else loop ");//No success message " +getmsg);
									  //log.log(LogStatus.INFO, ""+radio1.get(0).getAttribute("Name"));
								  }
								  
							}																				
			  }
				  break;
			  }
		}
		
		public void tslomsg() throws Exception
		{
			  String msg2 = "Success.\n TSLO reference number is:";
			  for(int m=0;m<2;)
			  {		
				  WebElement msg11 = w.findElementByClassName("SysListView32");		
				  List <WebElement> colcount1 = msg11.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.DataItem')]"));
				  System.out.println("item count "+colcount1.size());
				  
				  for(z=1;z<=colcount1.size();z++)
				  {		  													  
						  List<WebElement> radio1 = w.findElements(By.xpath("//*[contains(@Column, '"+z+"')]"));
						  System.out.println("radio1 size "+radio1.size());
						  
							for(int i=0;i<radio1.size();i++)
							{	
								getmsg = radio1.get(i).getAttribute("Name");
								System.out.println("in loop gm...."+getmsg);
								
								  if(getmsg.startsWith(msg2))
								  {																  
									  tsloremdot = getmsg.split("\\: ",0);
									  gettsloordno = tsloremdot[1].split("\\.", 0);
									  System.out.println("get on "+gettsloordno[0]);											  						  							  
									  
									  log.log(LogStatus.INFO, ""+getmsg);
									  log.log(LogStatus.INFO, "Order no " +gettsloordno[0]+ " placed successfully");
									  System.out.println("Completed successfully");	
									  
									  System.out.println("passed loop ");
									  storetsloordno = gettsloordno[0];
									  z=100;
									  break;																  
								  }
								  else
								  {
									  System.out.println("else loop ");//No success message " +getmsg);
									  log.log(LogStatus.INFO, ""+radio1.get(0).getAttribute("Name"));
								  }
								  
							}																				
			  }
				  break;
			  }
			  
		}
		
		public void normsmstatus() throws Exception
		{			
			log = report1.startTest("Placed order's status ");
			orderrpt();
			
			//count no of rows
			WebElement rpttable = w.findElement(By.name("Order Report [All]")).findElement(By.id("1003"));
			List < WebElement > rc3 = rpttable.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			System.out.println("row count "+ rc3.size());
							
			System.out.println("get sm ord no "+storesmordno);
			System.out.println("get ord no "+storenormordno);
			
			for(m=1;m<=rc3.size();m++)
			{
				capordno = rpttable.findElement(By.name("Row " + (m) + ", Column 4")).getText();
				System.out.println("ord no "+capordno);				  
													  
				///modify order						  
				getstat = rpttable.findElement(By.name("Row " + (m)+ ", Column 11")).getText();
				System.out.println("sm status "+getstat);
													  
				if(capordno.equals(storesmordno) && getstat.equals("TRAD"))
				{
					rpttable.findElement(By.name("Row " + (m)+ "")).click();
					log.log(LogStatus.INFO, "Select order no : "+capordno);
					log.log(LogStatus.INFO, "Super Multiple Order No " +capordno+ " is "+getstat);
					break;
				}
				else if(capordno.equals(storenormordno) && getstat.equals("TRAD"))
				{
					rpttable.findElement(By.name("Row " + (m)+ "")).click();
					log.log(LogStatus.INFO, "Select order no : "+capordno);
					log.log(LogStatus.INFO, "Normal Order No " +capordno+ " is " +getstat);
					break;
				}
				else //if(capordno.equals(getordno[0]) && getstat.equals("TRAD"))
					 {			
						//log.log(LogStatus.INFO, "Order No " +smordno+ " is  "+getstat);
					 }	
				}				
			  close();	
			
		}	
		
		public void tslostatus() throws Exception
		{
			tslorpt();
			
			//count no of rows
			WebElement rpttable = w.findElement(By.name("TSLO Order Report")).findElement(By.id("1426"));
			List < WebElement > rc3 = rpttable.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			System.out.println("row count "+ rc3.size());
			
			System.out.println("in loop get tslo ord no from status fn "+storetsloordno);
			
			if(TSLOchk.equals("ExistingTSLO"))
			{
					for(m=1;m<=rc3.size();m++)
					{	
						String getval = rpttable.findElement(By.name("Row " + (m) + ", Column 1")).getText();
						System.out.println("get val "+getval);
													
								if(getval.equals("TSLO Existing"))
								{										
									for(int n=1;n<20;n++)
									{	
										m=m+1;
										tsloordno = rpttable.findElement(By.name("Row " + (m) + ", Column 1")).getText();
										System.out.println("tslo ord no "+tsloordno);
										
										if(tsloordno.equals(storetsloordno))
										{										
												getstat = rpttable.findElement(By.name("Row " + (m)+ ", Column 13")).getText();
												System.out.println("tslo status "+getstat);
																					  
												if(tsloordno.equals(storetsloordno) && getstat.equals("TRAD"))
												{
													log.log(LogStatus.INFO, "TSLO Order No " +tsloordno+ " is " +getstat);	
												}
												else if((tsloordno.equals(storetsloordno) && getstat.equals("TRL")) || (tsloordno.equals(storetsloordno) && getstat.equals("WTRL")) )
												{	
													//log.log(LogStatus.INFO, "Order is " +getstat);
													vr.scrip = rpttable.findElement(By.name("Row " + (m) + ", Column 2")).getText();
													ME = rpttable.findElement(By.name("Row " + (m) + ", Column 3")).getText();
													vr.bprice = rpttable.findElement(By.name("Row " + (m) + ", Column 10")).getText();
													System.out.println("getscrip "+vr.scrip+ " me "+ME+ " sloprice "+vr.bprice);
													  
													close();
													mktdpth();																													 
												}
												else if(tsloordno.equals(storetsloordno) && getstat.equals("CAN"))
												{
													String remark = rpttable.findElement(By.name("Row " + (m)+ ", Column 14")).getText();
													System.out.println("remark "+remark);
												}
												else
												{
													log.log(LogStatus.INFO, "TSLO Order No " +tsloordno+ " is " +getstat);					
												}											
												break;												
											}
										}
									break;
								}
					}
			}
			
			else if(TSLOchk.equals("ExistingTSLO-BP"))
			{

				for(m=1;m<=rc3.size();m++)
				{	
					String getval = rpttable.findElement(By.name("Row " + (m) + ", Column 1")).getText();
					System.out.println("get val "+getval);
												
							if(getval.equals("Bracket Order Existing"))
							{										
								for(int n=1;n<20;n++)
								{	
									m=m+1;
									tsloordno = rpttable.findElement(By.name("Row " + (m) + ", Column 1")).getText();
									System.out.println("tslo ord no "+tsloordno);
									
									if(tsloordno.equals(storetsloordno))
									{										
											getstat = rpttable.findElement(By.name("Row " + (m)+ ", Column 13")).getText();
											System.out.println("tslo status "+getstat);
																				  
											if(tsloordno.equals(storetsloordno) && getstat.equals("TRAD"))
											{
												log.log(LogStatus.INFO, "TSLO Order No " +tsloordno+ " is " +getstat);	
											}
											else if((tsloordno.equals(storetsloordno) && getstat.equals("TRL")) || (tsloordno.equals(storetsloordno) && getstat.equals("WTRL")) )
											{	
												//log.log(LogStatus.INFO, "Order is "+getstat);
												vr.scrip = rpttable.findElement(By.name("Row " + (m) + ", Column 2")).getText();
												ME = rpttable.findElement(By.name("Row " + (m) + ", Column 3")).getText();
												vr.bprice = rpttable.findElement(By.name("Row " + (m) + ", Column 10")).getText();
												System.out.println("getscrip "+vr.scrip+ " me "+ME+ " sloprice "+vr.bprice);
												  
												close();
												mktdpth();	
											}
											else if(tsloordno.equals(storetsloordno) && getstat.equals("CAN"))
											{
												String remark = rpttable.findElement(By.name("Row " + (m)+ ", Column 14")).getText();
												System.out.println("remark "+remark);
											}
											else
											{
												log.log(LogStatus.INFO, "TSLO Order No " +tsloordno+ " is " +getstat);					
											}											
											break;												
										}
									}
								break;
							}
				}
		}
							
			else if(TSLOchk.equals("WLTSLO"))
			{
				for(m=1;m<=rc3.size();m++)
				{						
					System.out.println("m value "+m);
					String getval = rpttable.findElement(By.name("Row " + (m) + ", Column 1")).getText();
					System.out.println("get val "+getval);
					
							if(getval.equals("TSLO New"))
							{
								m=m+1;
								for(int i=1;i<rc3.size();i++)
								{
									//m=m+1;
									System.out.println("m for val "+m);
									System.out.println("i........ "+i);
									tsloordno = rpttable.findElement(By.name("Row " + (m) + ", Column 1")).getText();
									System.out.println("tslo ord no "+tsloordno);
									
									if(tsloordno.equals(storetsloordno))
									{
										for(int j=1;j<3;j++)
										{		
											System.out.println("j....m... "+m);
											System.out.println("j........ "+j);
											getstat = rpttable.findElement(By.name("Row " + (m)+ ", Column 13")).getText();
											System.out.println("tslo status "+getstat);
																				  
											if(tsloordno.equals(storetsloordno) && getstat.equals("TRAD"))
											{
												log.log(LogStatus.INFO, "TSLO Order No " +tsloordno+ " is " +getstat);					
											}
											else if((tsloordno.equals(storetsloordno) && getstat.equals("TRL")) || (tsloordno.equals(storetsloordno) && getstat.equals("WTRL")) )
											{	
												//log.log(LogStatus.INFO, "Order is "+getstat);
												vr.scrip = rpttable.findElement(By.name("Row " + (m) + ", Column 2")).getText();
												ME = rpttable.findElement(By.name("Row " + (m) + ", Column 3")).getText();
												vr.bprice = rpttable.findElement(By.name("Row " + (m) + ", Column 10")).getText();
												System.out.println("getscrip "+vr.scrip+ " me "+ME+ " sloprice "+vr.bprice);
												  
												close();
												mktdpth();															 
											}
											else if(tsloordno.equals(storetsloordno) && getstat.equals("CAN"))
											{
												String remark = rpttable.findElement(By.name("Row " + (m)+ ", Column 14")).getText();
												System.out.println("remark "+remark);
											}
											else
											{
												log.log(LogStatus.INFO, "TSLO Order No " +tsloordno+ " is " +getstat);					
											}
											m=m+1;
										}
										m=j;
										break;
									}
									m=m+2;									
								}
								break;
							}
				}
			}
							
			else if(TSLOchk.equals("WLTSLO-BP"))
			{
				for(m=1;m<=rc3.size();m++)
				{										
					System.out.println("m value "+m);
					String getval = rpttable.findElement(By.name("Row " + (m) + ", Column 1")).getText();
					System.out.println("get val "+getval);	
					
							if(getval.equals("Bracket Order New"))
							{
								m=m+1;
								for(int i=m;i<m+1;)
								{
									//m=m+1;
									System.out.println("m for val "+m);
									System.out.println("i........ "+i);
									tsloordno = rpttable.findElement(By.name("Row " + (m) + ", Column 1")).getText();
									System.out.println("tslo ord no "+tsloordno);
									
									if(tsloordno.equals(storetsloordno))
									{
										for(int j=1;j<3;j++)
										{		
											System.out.println("j....m... "+m);
											System.out.println("j........ "+j);
											getstat = rpttable.findElement(By.name("Row " + (m)+ ", Column 13")).getText();
											System.out.println("tslo status "+getstat);
																				  
											if(tsloordno.equals(storetsloordno) && getstat.equals("TRAD"))
											{
												log.log(LogStatus.INFO, "TSLO Order No " +tsloordno+ " is " +getstat);					
											}
											else if((tsloordno.equals(storetsloordno) && getstat.equals("TRL")) || (tsloordno.equals(storetsloordno) && getstat.equals("WTRL")) )
											{
												//log.log(LogStatus.INFO, "Order is "+getstat);
												vr.scrip = rpttable.findElement(By.name("Row " + (m) + ", Column 2")).getText();
												ME = rpttable.findElement(By.name("Row " + (m) + ", Column 3")).getText();
												vr.bprice = rpttable.findElement(By.name("Row " + (m) + ", Column 10")).getText();
												System.out.println("getscrip "+vr.scrip+ " me "+ME+ " sloprice "+vr.bprice);
												  
												close();
												mktdpth();																 
											}
											else if(tsloordno.equals(storetsloordno) && getstat.equals("CAN"))
											{
												String remark = rpttable.findElement(By.name("Row " + (m)+ ", Column 14")).getText();
												System.out.println("remark "+remark);
											}
											else
											{
												log.log(LogStatus.INFO, "TSLO Order No " +tsloordno+ " is " +getstat);					
											}
											m=m+1;
										}
										m=j;
										break;
									}
									m=m+2;
									
								}
								break;
							}
					}
				}	
			  close();				
		}
	  		
		public void mktdpth() throws Exception
		{		
			WebElement wltable = w.findElement(By.id("65280")).findElement(By.id("1000"));
			WebElement colcount = wltable.findElement(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));		
			List < WebElement > rc1 = wltable.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			System.out.println("row count "+ rc1.size());
						  
						  for(int k=1;k<=rc1.size();k++)
						  {
							  String getscrip = wltable.findElement(By.name("Row " + (k)+ ", Column 1")).getText();
							  String getme = wltable.findElement(By.name("Row " + (k) + ", Column 2")).getText();
							  System.out.println("getscrip "+getscrip+ " ME "+getme);
							  
							  if((getscrip.equals(vr.scrip) && getme.equals(ME)))// || (getscrip.equals(capscrip) && getme.equals(capME)))
							  {  
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
								  
								  for(int y=1;y<=2;y++)
								  {
									  System.out.println("y value "+y);				  
									  for(int z=1;z<txt.size();)
									  {														  
										  bpr = txt.get(z).getAttribute("Name");//.findElement(By.xpath("//*[contains(@Name), '"+msg2+"']")).getAttribute("Name");///
										  System.out.println("b pr "+bpr);
										  z=z+1;
										  bqty = "1";//txt.get(z).getAttribute("Name");
										  System.out.println("b qty "+bqty);
										  z=z+1;
										  bord = txt.get(z).getAttribute("Name");
										  System.out.println("b ord "+bord);
										  z=z+1;
										  spr = txt.get(z).getAttribute("Name");
										  System.out.println("s pr "+spr);
										  z=z+1;
										  sqty = "1";//txt.get(z).getAttribute("Name");
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
										  success();
										  
									  }
									  else
									  {
										  robot.keyPress(KeyEvent.VK_F1);
										  robot.keyRelease(KeyEvent.VK_F1);
										  w.findElement(By.id("1018")).sendKeys("1Y303");
										  w.findElement(By.id("1400")).sendKeys("1");
										  w.findElement(By.id("1404")).sendKeys(vr.bprice);											  
										  success();
									  }
										  
										  if(spr !=null &&  !spr.isEmpty())
										  {
											  
											  robot.keyPress(KeyEvent.VK_F1);
											  robot.keyRelease(KeyEvent.VK_F1);
											  w.findElement(By.id("1018")).sendKeys("1Y306");
											  w.findElement(By.id("1400")).sendKeys(sqty);
											  w.findElement(By.id("1404")).sendKeys("0");											  
											  success();	
											  										  
										  }
										  else
										  {
											  robot.keyPress(KeyEvent.VK_F2);
											  robot.keyRelease(KeyEvent.VK_F2);
											  w.findElement(By.id("1018")).sendKeys("1Y304");
											  w.findElement(By.id("1400")).sendKeys("1");
											  w.findElement(By.id("1404")).sendKeys(vr.bprice);											  
											  success();									  										  
										  }		
								  }
								  break;
						  }
							  else
							  {
								  System.out.println("scrip not found ");
							  }							  
						  }
						  Thread.sleep(1000);
						  robot.keyPress(KeyEvent.VK_ESCAPE);
						  Thread.sleep(500);								 
						  robot.keyRelease(KeyEvent.VK_ESCAPE);
						  Thread.sleep(500);
						  //log.log(LogStatus.INFO, "TRL Order is traded ");
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
				 
				 if(capscrip.equals(vr.scripchk) && capME.equals(getxchnge) && (delmark.equals("Super Multiple")))
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
						  log.log(LogStatus.INFO, "Select super multiple order scrip: "+capscrip);	
		  
						  ////check for square off order button
						  success();						  
						  log.log(LogStatus.INFO, "Order is squared off ");
						  break;									  
					  }								  
				 }
				 else
				 {
					 //log.log(LogStatus.INFO, "No super multiple order available");								 
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
						  ordermrk = vr.s.getRow(i).getCell(j).getStringCellValue();

						  j=j+1;
						  TSLOchk = vr.s.getRow(i).getCell(j).getStringCellValue();
						  
						  j=j+1;
						  otherclient = vr.s.getRow(i).getCell(j).getStringCellValue();
						  
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
						  								  
								  if(ordermrk.equals("Normal"))
								  {										  
									  if(TSLOchk.equals("WLTSLO"))
									  {
										  log = report1.startTest("Place TSLO Order for :"+vr.forreportscripname);
										  log.log(LogStatus.INFO, "Scrip " +vr.scrip+ " is selected");
										  WLTSLO();
									  }
									  else if(TSLOchk.equals("WLTSLO-BP"))
									  {
										  log = report1.startTest("Place Bracket TSLO Order for :"+vr.forreportscripname);
										  log.log(LogStatus.INFO, "Scrip " +vr.scrip+ " is selected");
										  WLTSLOBP();
									  }									  
									  else if(TSLOchk.equals("None"))
									  {
										  log = report1.startTest("Place Normal " +vr.scrip+ " Order");
										  log.log(LogStatus.INFO, "Scrip " +vr.scrip+ " is selected");
										  log.log(LogStatus.INFO, "Entered Client Code " +vr.clientcode);			
										  log.log(LogStatus.INFO, "Selected action is "+vr.action);			
										  log.log(LogStatus.INFO, "Entered quantity "+vr.qty);	
										  log.log(LogStatus.INFO, "Actual price is " +vr.getamt);
										  log.log(LogStatus.INFO, "Modified price is " +vr.roundoff);	
										  log.log(LogStatus.INFO, "Select Order Mark : " +ordermrk);
										  
										  if(vr.roundoff!=0)
										  {
											  success();
											  log.log(LogStatus.INFO, "Click on New "+vr.action+ " Order ");
											  
											  log.log(LogStatus.INFO, "Click on Cancel ");
											  
											  normmsg();
											  Normalorder();
										  }
										  else if(vr.roundoff==0)
										  {
											  if(vr.perc !=0)
											  {		
												  w.findElement(By.id("1342")).click();/////new buy order button
												  Thread.sleep(2000);
												  log.log(LogStatus.INFO, "Click on New "+vr.action+ " Order ");
												  			
												  robot.keyPress(KeyEvent.VK_ENTER);
												  Thread.sleep(2000);
												  
												  normmsg();
												  
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
												  w.findElement(By.id("1400")).sendKeys(String.valueOf(vr.qty));
												  log.log(LogStatus.INFO, "Entered quantity " +vr.qty);	
	
												  w.findElement(By.id("1404")).sendKeys("0.00"); ///for trade at mkt
												  log.log(LogStatus.INFO, "Place order at market price ");	
	
												  success();												  
											  }
											  else
											  {
												  success();
												  log.log(LogStatus.INFO, "Place order at market price ");
											  }
										  }
									  }	
								  }
								  else if(ordermrk.equals("Super Multiple"))
								  {	
									  log = report1.startTest("Place Super Multiple Order for :"+vr.forreportscripname);
									  log.log(LogStatus.INFO, "Scrip " +vr.scrip+ " is selected");
									  SuperMultiple();
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
			w.findElement(By.id("1421")).click();				
			w.findElement(By.id("1409")).sendKeys(String.valueOf(spread));
			log.log(LogStatus.INFO, "Spread is: " +spread);
			robot.keyPress(KeyEvent.VK_TAB);
			robot.keyRelease(KeyEvent.VK_TAB);
			gettrigpr = w.findElement(By.id("1411")).getText();
			System.out.println("tslo trig pr "+gettrigpr);
			
			success();
			log.log(LogStatus.INFO, "Place TSLO order from watchlist");	
			
			Thread.sleep(2000);
			tslomsg();
			
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
				  w.findElement(By.id("1018")).sendKeys(otherclient);
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
				  success();
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
			  w.findElement(By.id("1018")).sendKeys(roughclient);
			  w.findElement(By.id("1400")).sendKeys(String.valueOf(vr.qty)); 
			  w.findElement(By.id("1404")).sendKeys(String.valueOf(gettrigpr));	
			  success();
			  log.log(LogStatus.INFO, "Place normal order from watchlist to trade leg2 ");	
		}
		
		public void WLTSLOBP() throws Exception
		{			
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
			success();
			log.log(LogStatus.INFO, "Place TSLO order from watchlist with book profit ");	

			tslomsg();
			  
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
				  w.findElement(By.id("1018")).sendKeys(otherclient);
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
				  success();
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
			  w.findElement(By.id("1018")).sendKeys(roughclient);
			  w.findElement(By.id("1400")).sendKeys(String.valueOf(vr.qty)); 
			  w.findElement(By.id("1404")).sendKeys(String.valueOf(totbp));
			  success();
			  log.log(LogStatus.INFO, "Place normal order from watchlist with book profit to trade leg2 ");	
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
				 
				 //if(capscrip.equals(vr.getscrip) && capME.equals(getxchnge))
				 //{
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
							  //log.log(LogStatus.INFO, "Select scrip: "+capscrip);	
	
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
							  success();
							  log.log(LogStatus.INFO, "Place Existing TSLO order from open position ");	
							  
							  tslomsg();
							  
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
							  w.findElement(By.id("1018")).sendKeys(otherclient);
							  w.findElement(By.id("1401")).sendKeys(remdot[0]);
							  w.findElement(By.id("1399")).sendKeys(capME);
							  robot.keyPress(KeyEvent.VK_TAB);
							  robot.keyPress(KeyEvent.VK_TAB);
							  robot.keyPress(KeyEvent.VK_DELETE);
							  robot.keyRelease(KeyEvent.VK_TAB);
							  robot.keyRelease(KeyEvent.VK_DELETE);
							  w.findElement(By.id("1400")).sendKeys(getqty); 
							  w.findElement(By.id("1404")).sendKeys(String.valueOf(gettrigpr));	  
							  success();
							  log.log(LogStatus.INFO, "Place normal order from open position to trade leg2");	
							  break;									  
						  }								  
					 }
				 //}
				 else
				 {
					 //log.log(LogStatus.INFO, "No super multiple order available");								 
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
				 System.out.println("cap util "+capscrip);
				 remdot = capscrip.split("\\~",0);
				 System.out.println("cap util "+remdot[0]);

				 //capture ME for BSE
				 capME = abc.findElement(By.name("Row " + (p) + ", Column 2")).getText();
				 System.out.println("cap ME "+capME);
				 
				 delmark = abc.findElement(By.name("Row " + (p) + ", Column 16")).getText();
				 System.out.println("del mark "+delmark);
				 
				 nettrdval = abc.findElement(By.name("Row " + (p) + ", Column 9")).getText();
				 System.out.println("net trd val----- "+nettrdval);
				 
				 System.out.println("vrgetscrip "+vr.getscrip+ " getxchandge "+getxchnge);
				 
				 
				 //if(capscrip.equals(vr.getscrip) && capME.equals(getxchnge))
				 //{
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
							  //log.log(LogStatus.INFO, "Select scrip: "+capscrip);	
	  
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
							  success();
							  log.log(LogStatus.INFO, "Place Existing bracket TSLO order from open position ");	
							  
							  tslomsg();
							  
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
							  w.findElement(By.id("1018")).sendKeys(otherclient);
							  w.findElement(By.id("1401")).sendKeys(remdot[0]);
							  w.findElement(By.id("1399")).sendKeys(capME);
							  robot.keyPress(KeyEvent.VK_TAB);
							  robot.keyPress(KeyEvent.VK_TAB);
							  robot.keyPress(KeyEvent.VK_DELETE);
							  robot.keyRelease(KeyEvent.VK_TAB);
							  robot.keyRelease(KeyEvent.VK_DELETE);
							  w.findElement(By.id("1400")).sendKeys(getqty); 
							  w.findElement(By.id("1404")).sendKeys(String.valueOf(totbp));	  
							  success();
							  log.log(LogStatus.INFO, "Place normal order to trade leg2 ");	
							  break;									  
						  }								  
					 }
				 //}
				 else
				 {
					 //log.log(LogStatus.INFO, "No super multiple order available");								 
				 }
			}
			close();
		}
		
		public void Normalorder() throws Exception
		{
			  if(vr.roundoff!=0)
			  {	  				  
				  Normalmodify();
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
						
					success();
					log.log(LogStatus.INFO, "Order no " +getordno[0]+ " modified successfully");
					break;
				}
				else //if(capordno.equals(getordno[0]) && getstat.equals("TRAD"))
					 {			
						//log.log(LogStatus.INFO, "Order No " +capordno+ " is not OPN so cannot modify");
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
						success();
						log.log(LogStatus.INFO, "Order no " +getordno[0]+ " cancelled successfully");	
						break;
				  }	
				  else //if(capordno.equals(getordno[0]) && getstat.equals("TRAD"))
				  {
					  //log.log(LogStatus.INFO, "Order No " +capordno+ " is not OPN so cannot cancel");
				  }											
			  }
			  close();
		}
		
		public void SuperMultiple() throws Exception
		{
			w.findElement(By.id("1014")).click();
			
			if(vr.roundoff!=0)
			  {
				  w.findElement(By.name("Super Multiple")).click();
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
				  
				  //success();
				  w.findElement(By.id("1342")).click();/////new buy order button
				  Thread.sleep(3000);
				  w.findElement(By.id("1")).click();/////new buy order button
				  //robot.keyPress(KeyEvent.VK_ENTER);
				  Thread.sleep(3000);
				  w.findElement(By.name("Cancel")).click();
				  Thread.sleep(1500);
				  
				  log.log(LogStatus.INFO, "Place super multiple order ");
				  smmsg();
//////////////////////////////////			  
				  
			  }
			  else if(vr.roundoff==0)
			  {
				  for(int a=1;a<=1;) 
				  {
					  w.findElement(By.name("Cancel")).click();
					  
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
					  w.findElement(By.id("1018")).sendKeys(roughclient);
					  w.findElement(By.id("1400")).sendKeys(String.valueOf(addqty));
					  w.findElement(By.id("1014")).click();
					  w.findElement(By.name("Normal")).click();		
					  
					 success();
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
					  w.findElement(By.name("Super Multiple")).click();
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
					  w.findElement(By.id("1")).click();/////new buy order button
					  //robot.keyPress(KeyEvent.VK_ENTER);
					  Thread.sleep(3000);
					  w.findElement(By.name("Cancel")).click();
					  Thread.sleep(1500);
					  
					  log.log(LogStatus.INFO, "Place super multiple order ");					  
					  smmsg();					  
					  break;
				}
			  }
			
////////////////////////////////////////
			/*
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
			  w.findElement(By.id("1400")).sendKeys(String.valueOf(vr.qty));
			  w.findElement(By.id("1014")).click();
			  w.findElement(By.name("Super Multiple")).click();
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
			  Thread.sleep(3500);
			  robot.keyPress(KeyEvent.VK_ENTER);
			  Thread.sleep(3500);
			  w.findElement(By.name("Cancel")).click();
			  Thread.sleep(1500);
			  log.log(LogStatus.INFO, "Trade super multiple order ");	
			 */ 
////////////////////////////////////
			  ////assign scrip for smp
			  vr.scripchk = vr.getscrip;
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
			  
			  ////open watchlist
			  openWL();
			  
			  deltakesnapshot();
			  
			  //place order function
			  placeorder();
			  md();
			  smsquareoff();
			  normsmstatus();			  
			  tslostatus();  
			  
			  ///open all reports			  
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
			  Thread.sleep(15000);	
			  passtakesnapshot();
			  close();
			  
			  copy2pdf();			  
			  
			  exit();
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



