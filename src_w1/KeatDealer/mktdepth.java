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
import org.openqa.selenium.winium.WiniumDriver;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.LogStatus;

import Libraries.functions;
import Libraries.variables;
import Main.beforesuitedealer;

public class mktdepth extends beforesuitedealer
	{	
		WebElement keatpro,op;		
		public static int i,j,xcelrc,m,addqty,bpprice;
		Robot robot;
		public static String ME,bpr,sqty,bqty,spr,bord,sord,scrip;

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
		}
		
		public void success() throws Exception
		{
			  			  
			  w.findElement(By.id("1342")).click();/////new buy order button
			  //Thread.sleep(1500);
			  
			  WebDriverWait wait2= new WebDriverWait (w,500);
			  WebElement element2=wait2.until(ExpectedConditions.visibilityOfElementLocated(By.name("OK")));
			  element2.click();
			  //w.findElement(By.id("1")).click();/////success button
			  //Thread.sleep(1500);	  
			  
			  //cancel								  
			  w.findElement(By.id("2")).click();
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
							  
							  for(int y=1;y<=10;y++)
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
									  /*
									  w.findElement(By.id("1342")).click();/////new buy order button
									  Thread.sleep(1500);
									  
									  w.findElement(By.id("1")).click();/////success button
									  Thread.sleep(1500);							  
									  
									  //cancel								  
									  w.findElement(By.id("2")).click();
									  */
									  if(spr !=null &&  !spr.isEmpty())
									  {
										  robot.keyPress(KeyEvent.VK_F1);
										  robot.keyRelease(KeyEvent.VK_F1);
										  w.findElement(By.id("1018")).sendKeys("1Y306");
										  w.findElement(By.id("1400")).sendKeys(sqty);
										  w.findElement(By.id("1404")).sendKeys("0");
										  
										  success();
										  /*
										  w.findElement(By.id("1342")).click();/////new buy order button
										  Thread.sleep(1500);
										  
										  w.findElement(By.id("1")).click();/////success button
										  Thread.sleep(1500);							  
										  
										  //cancel								  
										  w.findElement(By.id("2")).click();
										  */
									  }
									  else
									  {
										  
										  robot.keyPress(KeyEvent.VK_F2);
										  robot.keyRelease(KeyEvent.VK_F2);
										  w.findElement(By.id("1018")).sendKeys("1Y302");
										  w.findElement(By.id("1400")).sendKeys(bqty);
										  w.findElement(By.id("1404")).sendKeys("0");
										  
										  success();
										  /*
										  w.findElement(By.id("1342")).click();/////new buy order button
										  Thread.sleep(1500);
										  
										  w.findElement(By.id("1")).click();/////success button
										  Thread.sleep(1500);							  
										  
										  //cancel								  
										  w.findElement(By.id("2")).click();
										  */
									  }
	
								  }
								  else
								  {
									  
									  robot.keyPress(KeyEvent.VK_F1);
									  robot.keyRelease(KeyEvent.VK_F1);
									  w.findElement(By.id("1018")).sendKeys("1Y305");
									  w.findElement(By.id("1400")).sendKeys(sqty);
									  w.findElement(By.id("1404")).sendKeys("0");
									  
									  success();
									  /*
									  w.findElement(By.id("1342")).click();/////new buy order button
									  Thread.sleep(1500);
									  
									  w.findElement(By.id("1")).click();/////success button
									  Thread.sleep(1500);							  
									  
									  //cancel								  
									  w.findElement(By.id("2")).click();
									  */
								  }		
								  Thread.sleep(500);
								  robot.keyPress(KeyEvent.VK_ESCAPE);
								  Thread.sleep(500);								 
								  robot.keyRelease(KeyEvent.VK_ESCAPE);
								  Thread.sleep(500);
							  }
						  }
						  
						  break;
					  }
				  }
			  }
			
		}
		
///////////////////////////////////////////////////////
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
							  
							  if(getscrip.equals(scrip) && getme.equals(ME))
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
								  
								  for(int y=1;y<=10;y++)
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
										  w.findElement(By.id("1404")).sendKeys("0");											  
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
											  w.findElement(By.id("1404")).sendKeys("0");											  
											  success();									  										  
										  }		
								  }
								  break;
						  }
							  else
							  {
								  System.out.println("scrip not found ");
							  }
							  Thread.sleep(500);
							  robot.keyPress(KeyEvent.VK_ESCAPE);
							  Thread.sleep(500);								 
							  robot.keyRelease(KeyEvent.VK_ESCAPE);
							  Thread.sleep(500);
						  }						  
		}		
///////////////////////////////////////////////////////		
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
			  md();

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



