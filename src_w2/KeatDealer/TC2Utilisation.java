package KeatDealer;

import java.awt.Robot;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.net.URL;
import java.text.DecimalFormat;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.winium.DesktopOptions;
import org.openqa.selenium.winium.WiniumDriver;
import org.testng.TestListenerAdapter;
import org.testng.TestNG;
import org.testng.annotations.Test;

import com.gargoylesoftware.htmlunit.javascript.host.event.MouseEvent;
import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import Libraries.functions;
import Libraries.variables;
import Main.beforesuitedealer;

public class TC2Utilisation extends beforesuitedealer
	{	
		public static String[] getordno,remdot;
		public static double cnvrtcaputil,placecnvrtcaputil,matchround;
		public static String caputilscrip,caputil,getmsg,capordno,getxchnge;
		public static String ME,capME;
		WebElement keatpro,op;		
		public static int i,j,xcelrc,m;
		Robot robot;
		
		variables vr = new variables();
		functions fn = new functions();
		
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
			  
			  vr.clientcode = vr.s.getRow(i).getCell(4).getStringCellValue();
			  System.out.println("client code "+vr.clientcode);
			  WebElement limittable = w.findElement(By.name("Net Position"));
			  limittable.findElement(By.id("1018")).sendKeys(vr.clientcode);
			  robot.keyPress(KeyEvent.VK_F5);
			  
			  //count no of rows
			  WebElement nptable = w.findElement(By.name("Net Position")).findElement(By.id("37033"));			  
			  List < WebElement > rc = nptable.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			  System.out.println("row count "+ rc.size());
			  
			  for(int k=1;k<rc.size();k++)
			  {				  
				  WebElement abc = nptable.findElement(By.name("Row " + (k) + ""));
				  caputilscrip = abc.findElement(By.name("Row " + (k) + ", Column 1")).getText();
				  System.out.println("cap util "+caputilscrip);
				  
				  //capture ME for BSE
				  capME = abc.findElement(By.name("Row " + (k) + ", Column 2")).getText();
				  System.out.println("cap ME "+capME);
				  
				  if(vr.scrip.equals(caputilscrip) && ME.equals(capME))
				  {
					  caputil = abc.findElement(By.name("Row " + (k) + ", Column 18")).getText();
					  System.out.println("cap util "+caputil);
					  cnvrtcaputil = Double.parseDouble(caputil);
					  break;
				  }
				  else
				  {
					  cnvrtcaputil = 0.0;
					  System.out.println("cap util "+cnvrtcaputil);
					  System.out.println("Scrip not available to capture position");
				  }	
				  vr.rw = vr.s.getRow(i);
				  vr.cell = vr.rw.createCell(variables.wputil);
				  vr.cell.setCellValue(cnvrtcaputil);
			  }			
			  		if(rc.size()==0)
			  		{
			  			cnvrtcaputil = 0.0;
			  			vr.rw = vr.s.getRow(i);///////////////updated
					  	vr.cell = vr.rw.createCell(variables.wputil);///////////////updated
					  	vr.cell.setCellValue(cnvrtcaputil);///////////////updated
			  		}			  					  		
		}
		
		public void openWL() throws Exception
		{				  
			FileInputStream f = new FileInputStream(srcfile);
			wb = new HSSFWorkbook(f);		
			vr.s = wb.getSheet("Utilisation");
			 
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
			w.findElement(By.id("1013")).click();
			WebElement openwltable = w.findElement(By.name("WatchList")).findElement(By.id("1013"));
			openwltable.findElement(By.xpath("//*[contains(@Name, '"+WL+"')]")).click();
			w.findElement(By.id("1")).click();
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
		
		public void netposition() throws Exception
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
			  
			  WebElement limittable1 = w.findElement(By.name("Net Position")).findElement(By.id("1011"));
			  WebElement colcountlt1 = limittable1.findElement(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));		
			  List < WebElement > rcnp1 = limittable1.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
			  System.out.println("row count "+ rcnp1.size());
			  
			  vr.totmar = limittable1.findElement(By.name("Row 1, Column 3")).getText();
			  System.out.println("tot mar "+vr.totmar);
			  log.log(LogStatus.INFO, "Total Margin is " +vr.totmar);
			  
			  for(int x=4;x<rcnp1.size();x++)
			  {
				  String gettotmargin = limittable1.findElement(By.name("Row " + (x) + ", Column 1")).getText();
				  System.out.println("get margin text "+gettotmargin);											  
				  
				  if(gettotmargin.equals("Margin Utilised (Cash&FNO)"))
				  {
					  	///click on margin utilised cashNfno
						vr.marginutilised = limittable1.findElement(By.name("Row " + (x) + ", Column 3")).getText();
						System.out.println(vr.marginutilised);
						log.log(LogStatus.INFO, "Margin Utilised is " +vr.marginutilised);						
						x=x+2;
				  }													  
				  if(gettotmargin.equals("Margin Available (Cash&FNO)"))
				  {
					  	///click on margin available cashNfno
					  	vr.marginavailable = limittable1.findElement(By.name("Row " + (x) + ", Column 3")).getText();
						System.out.println(vr.marginavailable);
						log.log(LogStatus.INFO, "Margin Available is " +vr.marginavailable);						
						break;
				  }	
			  }
			  
			//verify totalmargin - margin utilsed = margin available			  										
			Double ab = Double.parseDouble(vr.totmar)-Double.parseDouble(vr.marginutilised);
			BigDecimal bd = new BigDecimal(ab);
			System.out.println("bd " +bd);
			DecimalFormat df = new DecimalFormat("0.00");
			df.setMaximumFractionDigits(2);
			vr.sd = df.format(ab);
			System.out.println(vr.sd);			
						
			///check if above are equal
			if(vr.sd.equals(vr.marginavailable))
			{
				System.out.println("Equal");
				log.log(LogStatus.INFO, "Utilized -- As Required Margin and Margin Utilised are Equal");
				log.log(LogStatus.INFO, "Total Margin & Margin utilized are equal to Margin Available " +vr.sd);
													
				vr.rw = vr.s.getRow(i);
				vr.cell = vr.rw.createCell(variables.wstatus);
				vr.cell.setCellValue("Pass");
			}
			else
			{
				log.log(LogStatus.ERROR, "Total Margin & Margin utilized are not equal to Margin Available "+vr.sd);
				
				vr.rw = vr.s.getRow(i);
				vr.cell = vr.rw.createCell(variables.wstatus);
				vr.cell.setCellValue("Fail");
			}
			
			close();
		}
		
		public void placeorder() throws Exception
		{			  		
	  		FileInputStream f = new FileInputStream(srcfile);
			wb = new HSSFWorkbook(f);		
			vr.s = wb.getSheet("Utilisation");
			
			//String getmsg = msg.findElement(By.xpath("//*[contains(@Name, '"+scrip2+"')]")).getAttribute("Name");///working

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
						  log = report1.startTest("Place Order");
						  log.log(LogStatus.INFO, "Watchlist opened");			
						  
						  vr.scrip = vr.s.getRow(i).getCell(j).getStringCellValue();
						  System.out.println("scrip "+vr.scrip);
						  log.log(LogStatus.INFO, "Scrip " +vr.scrip+ " is selected");			
						  
						  ///excel for BSE
						  j=j+1;
						  ME = vr.s.getRow(i).getCell(j).getStringCellValue();
						  System.out.println("ME "+ME);
						  
						  ///capture util  
						  captureutil();
						  close();
						  
						  for(int k=1;k<=rc1.size();k++)
						  {
							  vr.getscrip = wltable.findElement(By.name("Row " + (k) + ", Column 1")).getText();
							  System.out.println("get scrip " +vr.getscrip);
							  wltable.findElement(By.name("Row " + (k)+ ", Column 1")).click();
							  getxchnge = wltable.findElement(By.name("Row " + (k)+ ", Column 2")).getText();
							  System.out.println("get exchnage "+getxchnge);
							  vr.getprice = wltable.findElement(By.name("Row " + (k)+ ", Column 9")).getText();
							  System.out.println("get price "+vr.getprice);
							  double chngegetprice = Double.parseDouble(vr.getprice); 

							  if(vr.scrip.equals(vr.getscrip) && ME.equals(getxchnge))
							  {								  
								  robot.keyPress(KeyEvent.VK_F1);
								  
								  ///select action
								  j=j+1;
								  vr.action = vr.s.getRow(i).getCell(j).getStringCellValue();

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
								  w.findElement(By.id("1018")).sendKeys(vr.clientcode);
								  log.log(LogStatus.INFO, "Entered Client Code " +vr.clientcode);			
								  log.log(LogStatus.INFO, "Selected action is "+vr.action);			
								  
								  robot.keyPress(KeyEvent.VK_TAB);
								  robot.keyPress(KeyEvent.VK_DELETE);	
								  
								  ///enter qty
								  j=j+1;
								  vr.qty = (int) vr.s.getRow(i).getCell(j).getNumericCellValue();
								  w.findElement(By.id("1400")).sendKeys(String.valueOf(vr.qty));
								  log.log(LogStatus.INFO, "Entered quantity "+vr.qty);			
								  
								  //read multiple from excel
								  j=j+1;
								  vr.multiple = (float) vr.s.getRow(i).getCell(j).getNumericCellValue();
								  System.out.println("multiple "+vr.multiple);
								  
								  ///capture actual price
								  robot.keyPress(KeyEvent.VK_TAB);
								  vr.getamt = w.findElement(By.id("1404")).getText();
								  System.out.println("getamount "+vr.getamt);							  
								  								  
								  ///get price perc
								  j=j+1;
								  vr.perc = (float) vr.s.getRow(i).getCell(j).getNumericCellValue();
								  vr.getamt1 = (Double.parseDouble(vr.getamt)*vr.perc)/100;
								  double chngestrnggetamt = Double.parseDouble(vr.getamt);
								  System.out.println("vr.action "+vr.action);
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
								  log.log(LogStatus.INFO, "Modified price is " +vr.roundoff);	
								  								  								  
								  robot.keyPress(KeyEvent.VK_DELETE);
								  
								  if(vr.roundoff<=0)
								  {
									  vr.roundoff=0.0;
									  ///calculate reqd margin
									  if(vr.scrip.equals(vr.getscrip) && getxchnge.equals("NSE-FX"))
									  {
										  Double calcutil = (chngegetprice * vr.qty*1000)/vr.multiple;
										  vr.newroundoff = (Math.round(calcutil*100.0)/100.0);
										  System.out.println("calc util "+vr.newroundoff);			  
										  log.log(LogStatus.INFO, "Required Margin is : "+vr.newroundoff);
									  }
									  else
									  {
										  Double calcutil = (chngegetprice * vr.qty)/vr.multiple;
										  vr.newroundoff = (Math.round(calcutil*100.0)/100.0);
										  System.out.println("calc util "+vr.newroundoff);			  
										  log.log(LogStatus.INFO, "Required Margin is : "+vr.newroundoff);	
									  }
								  }	
								  else
								  {
									  if(vr.scrip.equals(vr.getscrip) && getxchnge.equals("NSE-FX"))
									  {
										  ///calculate reqd margin
										  Double calcutil = (vr.roundoff * vr.qty*1000)/vr.multiple;
										  vr.newroundoff = (Math.round(calcutil*100.0)/100.0);
										  System.out.println("calc util "+vr.newroundoff);			  
										  log.log(LogStatus.INFO, "Required Margin is : "+vr.newroundoff);
									  }
									  else
									  {
										  ///calculate reqd margin
										  Double calcutil = (vr.roundoff * vr.qty)/vr.multiple;
										  vr.newroundoff = (Math.round(calcutil*100.0)/100.0);
										  System.out.println("calc util "+vr.newroundoff);			  
										  log.log(LogStatus.INFO, "Required Margin is : "+vr.newroundoff);
									  }
									  
								  }
								  w.findElement(By.id("1404")).sendKeys(String.valueOf(vr.roundoff));								  							  
								  
								  j=j+1;
								  vr.rw = vr.s.getRow(i);
								  vr.cell = vr.rw.createCell(j);
								  vr.cell.setCellValue(vr.getamt);
								  
								  j=j+1;
								  vr.rw = vr.s.getRow(i);
								  vr.cell = vr.rw.createCell(j);
								  vr.cell.setCellValue(vr.roundoff);
								  
								  ///chk for AMO
								  j=j+1;
								  String amochk = vr.s.getRow(i).getCell(j).getStringCellValue();
								  System.out.println("amochk "+amochk);
								  if(amochk.equals("YES"))
								  {
									  w.findElement(By.id("1420")).click();
									  log.log(LogStatus.INFO, "Select Place AMO checkbox");
								  }
									
								  w.findElement(By.id("1342")).click();/////new buy order button
								  Thread.sleep(1500);
								  log.log(LogStatus.INFO, "Click on New "+vr.action+ " Order ");			
								  
								  //confirmation enter
								  robot.keyPress(KeyEvent.VK_ENTER);
								  Thread.sleep(1500);
								  //success enter
								  robot.keyPress(KeyEvent.VK_ENTER);
								  Thread.sleep(1000);
								  //cancel								  
								  w.findElement(By.id("2")).click();
								  log.log(LogStatus.INFO, "Click on Cancel ");			
								  
								  	  String scrip2 = "Success. ABWS Order number is:";
									  for(int m=0;m<4;)
									  {					  				  
										  WebElement msg = w.findElementByClassName("SysListView32");
										  getmsg = msg.findElement(By.xpath("//*[contains(@Name, '"+scrip2+"')]")).getAttribute("Name");///working
										  System.out.println("get message "+getmsg);
										  if(getmsg.startsWith(scrip2))
										  {
											  remdot = getmsg.split("\\:",0);
											  getordno = remdot[1].split("\\.", 0);
											  System.out.println("get on "+getordno[0]);											  						  							  
											  
											  log.log(LogStatus.INFO, ""+getmsg);
											  log.log(LogStatus.INFO, "Order no " +getordno[0]+ " placed successfully");
											  System.out.println("Completed successfully");
											  
											  vr.rw = vr.s.getRow(i);
											  vr.cell = vr.rw.createCell(variables.wpreqmar);
											  vr.cell.setCellValue(vr.newroundoff);
											  
											  Double capcaputil = cnvrtcaputil;
											  System.out.println("capture cap util "+capcaputil);
											  captureutil();
											  //close();
											  Double match = cnvrtcaputil - capcaputil;
											  System.out.println("match "+match);
											  vr.matchround = Math.round(match*100.0)/100.0;
											  System.out.println("f amt" +vr.matchround);
											  log.log(LogStatus.INFO, "Margin Utilised for Cash " +vr.matchround);
											  
											  vr.rw = vr.s.getRow(i);
											  vr.cell = vr.rw.createCell(variables.wmutil);
											  vr.cell.setCellValue(cnvrtcaputil);
													
											  vr.rw = vr.s.getRow(i);
											  vr.cell = vr.rw.createCell(variables.wpmarutil);
											  vr.cell.setCellValue(vr.matchround);
								  
										  	////to check if reqd margin and captured util amt are same
											if(vr.matchround == vr.newroundoff)
											{					
												///open limits
												  Thread.sleep(500);
												  robot.keyPress(KeyEvent.VK_SHIFT);
												  Thread.sleep(500);
												  robot.keyPress(KeyEvent.VK_L);
												  Thread.sleep(500);
												  robot.keyRelease(KeyEvent.VK_SHIFT);
												  Thread.sleep(500);
												  
												  netposition();
												  
												  vr.rw = vr.s.getRow(i);
												  vr.cell = vr.rw.createCell(variables.wptotmar);
												  vr.cell.setCellValue(vr.totmar);
												  
												  vr.rw = vr.s.getRow(i);
												  vr.cell = vr.rw.createCell(variables.wptotmarutil);
												  vr.cell.setCellValue(vr.marginutilised);
												  
												  vr.rw = vr.s.getRow(i);
												  vr.cell = vr.rw.createCell(variables.wpmaravail);
												  vr.cell.setCellValue(vr.marginavailable);
													
												  vr.rw = vr.s.getRow(i);
												  vr.cell = vr.rw.createCell(variables.wpactmaravail);
												  vr.cell.setCellValue(vr.sd);
												  
												  vr.rw = vr.s.getRow(i);
												  vr.cell = vr.rw.createCell(variables.wstatus);
												  vr.cell.setCellValue("Pass");
											}
											else
											{
												close();
												log.log(LogStatus.ERROR, "Not Utilized -- As Required Margin and Margin Utilised are Unequal");
												log.log(LogStatus.ERROR, "Total Margin & Margin utilized are not equal to Margin Available ");
												
												vr.rw = vr.s.getRow(i);
												vr.cell = vr.rw.createCell(variables.wstatus);
												vr.cell.setCellValue("Fail");
											}
												break;
										  }
										  else
										  {
											  log.log(LogStatus.INFO, "Order No not generated... ");
											  System.out.println("No success message ");
										  }
									  }
								  break;
							  }
							  else
							  {
								  System.out.println("scrip not found ");
							  }
					  
				  			}
						  
						  ///////////////manage order
							log = report1.startTest("Change Order");
							robot.keyPress(KeyEvent.VK_ALT);
							Thread.sleep(500);
							robot.keyPress(KeyEvent.VK_R);
							Thread.sleep(500);
							robot.keyPress(KeyEvent.VK_O);
							Thread.sleep(1000);
							robot.keyRelease(KeyEvent.VK_ALT);
							log.log(LogStatus.INFO, "Open Order Report ");
							
							//enter client code
							WebElement rpttable1 = w.findElement(By.name("Order Report [All]"));//.findElement(By.id("1003"));
							rpttable1.findElement(By.id("1018")).sendKeys(vr.clientcode);
							robot.keyPress(KeyEvent.VK_F5);
							
							//count no of rows
							WebElement rpttable = w.findElement(By.name("Order Report [All]")).findElement(By.id("1003"));
							List < WebElement > rc3 = rpttable.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
							System.out.println("row count "+ rc3.size());
							
							  for(m=1;m<=rc3.size();m++)
							  {
								  capordno = rpttable.findElement(By.name("Row " + (m) + ", Column 4")).getText();
								  System.out.println("ord no "+capordno);				  
								  
								  ///modify order						  
								  String getstat = rpttable.findElement(By.name("Row " + (m)+ ", Column 11")).getText();
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
									  
									  Double newcap = vr.matchround;
									  System.out.println("newcap " +newcap);	
									  
									  System.out.println("value of J "+j);
									  j=j+1;
									  System.out.println("value of i "+i);
									  vr.modqty = (int) vr.s.getRow(i).getCell(j).getNumericCellValue();
									  w.findElement(By.id("1400")).sendKeys(String.valueOf(vr.modqty));
									  log.log(LogStatus.INFO, "Modified quantity : "+vr.modqty);									  							  
									  
									  ///calc modified qty
									  if(vr.scrip.equals(vr.getscrip) && getxchnge.equals("NSE-FX"))
									  {
										  ///calculate mod reqd margin
										  vr.MRM = ((1000*vr.modqty * vr.roundoff ) / vr.multiple);
										  System.out.println("MRM " +vr.MRM);
										  vr.roundoff1 = (Math.round(vr.MRM*100.0)/100.0);
										  System.out.println("Modified Reqd margin is " +vr.roundoff1);
										  //log.log(LogStatus.INFO, "Modified Required Margin is : "+vr.roundoff1);
									  }
									  else
									  {
										  vr.MRM = ((vr.modqty * vr.roundoff ) / vr.multiple);
										  System.out.println("MRM " +vr.MRM);
										  vr.roundoff1 = (Math.round(vr.MRM*100.0)/100.0);
										  System.out.println("Modified Reqd margin is " +vr.roundoff1);
										  //log.log(LogStatus.INFO, "Modified Required Margin is : "+vr.roundoff1);
									  }								  
									  
									  vr.rw = vr.s.getRow(i);
									  vr.cell = vr.rw.createCell(variables.wmreqmar);
									  vr.cell.setCellValue(vr.roundoff1);
									  
									  w.findElement(By.id("1342")).click();
									  Thread.sleep(1000);			  			  
									  w.findElement(By.id("6")).click();
									  Thread.sleep(1000);
									  w.findElement(By.id("1")).click();
									  Thread.sleep(1000);
									  w.findElement(By.id("2")).click();	///cancel modify dialog box		
									  log.log(LogStatus.INFO, "Order no " +getordno[0]+ " modified successfully");
									  log.log(LogStatus.INFO, "Modified Required Margin is : "+vr.roundoff1);
									  
									  ///chk mod reqd margin
									  close();
									  Thread.sleep(500);
									  robot.keyPress(KeyEvent.VK_ALT);
									  Thread.sleep(500);
									  robot.keyPress(KeyEvent.VK_R);
									  Thread.sleep(500);
									  robot.keyPress(KeyEvent.VK_P);
									  Thread.sleep(1000);
									  robot.keyRelease(KeyEvent.VK_ALT);									  
									  
									  //enter client code
									  WebElement nptable2 = w.findElement(By.name("Net Position"));
									  nptable2.findElement(By.id("1018")).sendKeys(vr.clientcode);
									  robot.keyPress(KeyEvent.VK_F5);
									  
									  //count no of rows
									  WebElement nptable1 = w.findElement(By.name("Net Position")).findElement(By.id("37033"));			  
									  List < WebElement > rc2 = nptable1.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
									  System.out.println("row count "+ rc2.size());
									  
									  ///////////open position
									  for(int a=1;a<rc2.size();a++)
									  {
										  String placecaputilscrip = nptable1.findElement(By.name("Row " + (a) + ", Column 1")).getText();
										  System.out.println("place cap util scrip "+placecaputilscrip);
										  
										  //for BSE
										  String getme = nptable1.findElement(By.name("Row " + (a) + ", Column 2")).getText();
										  System.out.println("get me "+getme);
										  
										  if(vr.scrip.equals(placecaputilscrip) && ME.equals(getme))
										  {
											  String placecaputil = nptable1.findElement(By.name("Row " + (a) + ", Column 18")).getText();
											  System.out.println("cap new util "+placecaputil);
											  placecnvrtcaputil = Double.parseDouble(placecaputil);
											  System.out.println("cnvrt cap util again "+cnvrtcaputil);
											  Double match = placecnvrtcaputil - cnvrtcaputil;
											  System.out.println("minus " +match);
											  vr.modmatchround = Math.round(match*100.0)/100.0;
											  System.out.println("round minus " +vr.modmatchround);						
											  //log.log(LogStatus.INFO, "Margin utilised is : " +vr.modmatchround);
											  
											  vr.rw = vr.s.getRow(i);
											  vr.cell = vr.rw.createCell(variables.wcutil);
											  vr.cell.setCellValue(placecnvrtcaputil);
											  
											  if(vr.roundoff1 == vr.modmatchround)
											  {
												///open limits
												  Thread.sleep(500);
												  robot.keyPress(KeyEvent.VK_SHIFT);
												  Thread.sleep(500);
												  robot.keyPress(KeyEvent.VK_L);
												  Thread.sleep(500);
												  robot.keyRelease(KeyEvent.VK_SHIFT);
												  Thread.sleep(500);
												  
												  netposition();
												  
												  vr.rw = vr.s.getRow(i);
												  vr.cell = vr.rw.createCell(variables.wmtotmar);
												  vr.cell.setCellValue(vr.totmar);
												  
												  vr.rw = vr.s.getRow(i);
												  vr.cell = vr.rw.createCell(variables.wmtotmarutil);
												  vr.cell.setCellValue(vr.marginutilised);
												  
												  vr.rw = vr.s.getRow(i);
												  vr.cell = vr.rw.createCell(variables.wmmaravail);
												  vr.cell.setCellValue(vr.marginavailable);
													
												  vr.rw = vr.s.getRow(i);
												  vr.cell = vr.rw.createCell(variables.wmactmaravail);
												  vr.cell.setCellValue(vr.sd);
											  }
											  else
											  {
												  close();
												  //log.log(LogStatus.INFO, "Not Utilised");
												  System.out.println("utilisation is unequal ");
											  }
											   break;
										  }
										  else
										  {
											  System.out.println("Scrip not available to capture position");
										  }
									  }
									  break;
								  }
								  else if(capordno.equals(getordno[0]) && getstat.equals("TRAD"))
								  {									  
									  log.log(LogStatus.INFO, "Order No " +capordno+ " is already TRADED so cannot modify");
									  System.out.println("mod else.........");
									  close();
									  break;									  
								  }
								  else 
								  {
									  log.log(LogStatus.INFO, "Order No " +capordno+ " Status is " +getstat+ " so cannot modify");
									  close();
									  break;
								  }
							  }
							  					
						/////////cancel
								//open order report
							  		//close();
									log = report1.startTest("Cancel Order");
									robot.keyPress(KeyEvent.VK_ALT);
									Thread.sleep(500);
									robot.keyPress(KeyEvent.VK_R);
									Thread.sleep(500);
									robot.keyPress(KeyEvent.VK_O);
									Thread.sleep(500);
									robot.keyRelease(KeyEvent.VK_ALT);
									Thread.sleep(500);
											  
									//enter client code
									WebElement rpttable2 = w.findElement(By.name("Order Report [All]"));//.findElement(By.id("1003"));
									rpttable2.findElement(By.id("1018")).sendKeys(vr.clientcode);
									robot.keyPress(KeyEvent.VK_F5);
									
									  WebElement rpttablecan = w.findElement(By.name("Order Report [All]")).findElement(By.id("1003"));
									  List < WebElement > rc4 = rpttablecan.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
									  System.out.println("row count "+ rc4.size());
									  
									  for(int c=1;c<=rc4.size();c++)
									  {
									  	 capordno = rpttablecan.findElement(By.name("Row " + (c) + ", Column 4")).getText();
										 System.out.println("ord no "+capordno);				  
											  
										 ///cancel order						  
										 String getstat1 = rpttablecan.findElement(By.name("Row " + (c)+ ", Column 11")).getText();
										 System.out.println("can status "+getstat1);
										 
									  if(capordno.equals(getordno[0]) && getstat1.equals("OPN") || getstat1.equals("AMO"))
									  {
										  	rpttablecan.findElement(By.name("Row " + (m)+ "")).click();
											log.log(LogStatus.INFO, "Select order no : "+capordno);
											Thread.sleep(500);
											robot.mousePress(InputEvent.BUTTON3_DOWN_MASK);
											Thread.sleep(500);
											robot.mouseRelease(InputEvent.BUTTON3_DOWN_MASK);
											Thread.sleep(500);
											robot.keyPress(KeyEvent.VK_C);
											Thread.sleep(500);
											w.findElement(By.id("1342")).click();
											Thread.sleep(1000);			  			  
											w.findElement(By.id("6")).click();
											Thread.sleep(1000);
											w.findElement(By.id("1")).click();
											Thread.sleep(1000);
											w.findElement(By.id("2")).click();	///cancel cancel dialog box
											log.log(LogStatus.INFO, "Order no " +getordno[0]+ " cancelled successfully");	
											 
											vr.newamt = cnvrtcaputil;
											vr.rw = vr.s.getRow(i);
											vr.cell = vr.rw.createCell(variables.wclastutil);
											vr.cell.setCellValue(vr.newamt);
											 
											///call capture total										
											captureutil();
											Double canamt = vr.newamt - cnvrtcaputil;
											System.out.println("can amt " +Math.abs(canamt));
											log.log(LogStatus.INFO, "Cancelled Required Margin is " +Math.abs(canamt));									
											
											///open limits
											robot.keyPress(KeyEvent.VK_ALT);
											Thread.sleep(500);
											robot.keyPress(KeyEvent.VK_R);
											Thread.sleep(500);
											robot.keyPress(KeyEvent.VK_P);
											Thread.sleep(1000);
											robot.keyRelease(KeyEvent.VK_ALT);
											Thread.sleep(500);									
											///open limits
											Thread.sleep(500);
											robot.keyPress(KeyEvent.VK_SHIFT);
											Thread.sleep(500);
											robot.keyPress(KeyEvent.VK_L);
											Thread.sleep(500);
											robot.keyRelease(KeyEvent.VK_SHIFT);
											Thread.sleep(500);
											  
											netposition();								 
											
											vr.rw = vr.s.getRow(i);
											vr.cell = vr.rw.createCell(variables.wctotmar);
											vr.cell.setCellValue(vr.totmar);
											
											vr.rw = vr.s.getRow(i);
											vr.cell = vr.rw.createCell(variables.wctotmarutil);
											vr.cell.setCellValue(vr.marginutilised);
											
											vr.rw = vr.s.getRow(i);
											vr.cell = vr.rw.createCell(variables.wcmaravail);
											vr.cell.setCellValue(vr.marginavailable);
											
											vr.rw = vr.s.getRow(i);
											vr.cell = vr.rw.createCell(variables.wcactmaravail);
											vr.cell.setCellValue(vr.sd);
											
											close();											
											break;
									  }
									  else if(capordno.equals(getordno[0]) && getstat1.equals("TRAD"))
									  {
										  log.log(LogStatus.INFO, "Order No " +capordno+ " is already TRADED so cannot cancel");
										  System.out.println("can else.........");
										  close();
										  break;
									  }
									  else
									  {
										  log.log(LogStatus.INFO, "Order No " +capordno+ " Status is " +getstat1+ " so cannot cancel");
										  close();
										  break;
									  }
								  }								  
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////									  
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
		
		
		 @Test
		 public void placemodcanorder() throws IOException, Exception
		 {		  		  	
			 			  			  
			  ///report creation
			  report1 = new ExtentReports("D:\\testfile\\winapputilisation.html");
			  log = report1.startTest("Login Details");
			  	
			  ////call login function
			  Login();
			  
			  robot = new Robot();	
			  
			  ////open watchlist
			  openWL();
			  //place order function
			  placeorder();	
			  			  			  			  
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
				  //w.get("D:\\testfile\\winapputilisation.html");				  
				  //w.get(srcfile);
			  }
		 
		}

