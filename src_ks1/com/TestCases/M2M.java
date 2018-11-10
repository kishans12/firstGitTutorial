package com.TestCases;

import java.awt.Robot;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.testng.TestListenerAdapter;
import org.testng.TestNG;
import org.testng.annotations.Test;

import com.library.functions;
import com.library.variables;
import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.LogStatus;

import Main.BSOrderXL;
import Utilisation.ManageWatchlist;

public class M2M extends BSOrderXL
{
	public static HSSFSheet s;
	public static String popup1;
	public static String value;
	public static int k,rc;
	public static WebElement ele;
	public static String getexactexpdate;
	
	public static int i,j,xcelrc,m;
	public static String[] getordno,remdot;
	public static String capscrip,capLTP,getmsg,capordno,getxchnge,getaction;
	public static String ME,capME,capqty,capnettrdval,capunrealpl;
	public static String stclientcode,xcelordno,prevLTP,price;
	public double calc;
	
	com.library.functions fn = new com.library.functions();
	com.library.variables vr = new com.library.variables();
				
	public void manage() throws Exception
	{							
		///click funds >>equity/derivatives
		Actions act = new Actions(w);
		act.moveToElement(w.findElement(By.linkText("Watchlist"))).click().perform();	
		WebElement Sublink = w.findElement(By.linkText("Manage"));
		Sublink.click();
		log.log(LogStatus.INFO, "Click on Watchlist >> Manage");
		
		///click Add
		w.findElement(By.partialLinkText("ADD")).click();
		log.log(LogStatus.INFO, "Click on Add");
		
		//enter watchlist name
		vr.watchlistname = s.getRow(1).getCell(5).getStringCellValue();  	
		w.findElement(By.xpath("//*[@id=\'container\']/div/form/table[1]/tbody/tr/td[2]/input")).sendKeys(vr.watchlistname);
		log.log(LogStatus.INFO, "Enter Watchlist Name : " +vr.watchlistname);
		
		///click submit
		w.findElement(By.partialLinkText("Submit")).click();
		log.log(LogStatus.INFO, "Click on Submit");
		
		///to check if watchlistname already exist
		String check = "Error";
		if(w.getPageSource().contains(check))
		{
			String catcherror = w.findElement(By.xpath("//*[@id=\'container\']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
			log.log(LogStatus.INFO, "" +catcherror);
			w.findElement(By.partialLinkText("BACK")).click();
			
			///go to manage
			Actions act1 = new Actions(w);
			act1.moveToElement(w.findElement(By.linkText("Watchlist"))).click().perform();	
			WebElement Sublink1 = w.findElement(By.linkText("Manage"));
			Sublink1.click();
			
			///check and delete same name
			if(w.getPageSource().contains(vr.watchlistname))
			{
				List<WebElement> radio = w.findElements(By.xpath("//input[@name='id' and @type = 'radio']"));
				for(int i=0;i<radio.size();i++)
				{
					ele = radio.get(i);
					String value=ele.getAttribute("id");
					vr.chk = w.findElement(By.xpath(".//td[input[@id ='"+value+"']]")).getText();						
					
					if(vr.chk.contains(vr.watchlistname))
					{
						ele.click();
						break;
					}
				}
				
				w.findElement(By.partialLinkText("DELETE")).click();
				log.log(LogStatus.INFO, "Existed Watchlist deleted successfully");
				w.findElement(By.partialLinkText("ADD")).click();
				log.log(LogStatus.INFO, "Add watchlist");
				w.findElement(By.xpath("//*[@id=\'container\']/div/form/table[1]/tbody/tr/td[2]/input")).sendKeys(vr.watchlistname);
				log.log(LogStatus.INFO, "Enter Watchlist Name : " +vr.watchlistname);
				w.findElement(By.partialLinkText("Submit")).click();
				log.log(LogStatus.INFO, "Click on Submit");
				w.findElement(By.xpath("//div[@id='container']//a[contains(text(), '"+vr.watchlistname+"')]")).click();
				log.log(LogStatus.INFO, "Select " +vr.watchlistname+ " to add scrips");
				
			}
		}
		else
			{
				///select added watchlist
				w.findElement(By.xpath("//div[@id='container']//a[contains(text(), '"+vr.watchlistname+"')]")).click();
				log.log(LogStatus.INFO, "Watchlist " +vr.watchlistname+ " gets added successfully " );
				
			}				
	}
	
	public void addloopscrips() throws Exception
	{				
		int i;
		int j;
		
		rc = s.getLastRowNum() - s.getFirstRowNum();
	  	System.out.println("rc "+rc);
	  	  	
	  	for(i=5;i<rc+1;i++)
	  	{
	  		Row row = s.getRow(i);
	  		for(j=0;j<row.getLastCellNum();)
	  		{
	  			String flag = s.getRow(i).getCell(j).getStringCellValue();
	  			
	  			if(flag.equals("Y"))
	  			{
		  			w.findElement(By.partialLinkText("ADD")).click();
		  			
		  			j=j+2;
		  			vr.scrip = s.getRow(i).getCell(j).getStringCellValue();
		  			w.findElement(By.id("search_list")).sendKeys(vr.scrip);
		  			
		  			j=j+1;
		  			String ME = s.getRow(i).getCell(j).getStringCellValue();
		  			w.findElement(By.id("me")).sendKeys(ME);
		  				  				  			
		  			j=j+1;
		  			String IT = s.getRow(i).getCell(j).getStringCellValue();
		  			w.findElement(By.id("inst_type")).sendKeys(IT);
		  				  			
		  			w.findElement(By.partialLinkText("SUBMIT")).click();	  								  						  								
	
		  			//to check if scrip exist
		  			if(w.getPageSource().contains(vr.scrip.toUpperCase()))
		        	{	  		
				        fnodate = s.getRow(8).getCell(2).getStringCellValue();
				        currdate = s.getRow(8).getCell(3).getStringCellValue();
		  				fn.getdate();		  				
		        	}
		  			else
		  			{
		  				w.findElement(By.partialLinkText("CANCEL")).click();
			  			log.log(LogStatus.INFO, "No matching stocks found for the Scrip " +vr.scrip.toUpperCase());
			  			
			  			j=j+1;
			  			vr.rw = s.getRow(i);
			  			vr.cell = vr.rw.createCell(j);
			  			vr.cell.setCellValue("null");
			  			
			  			j=j+1;		  			  			
			  			vr.rw = s.getRow(i);
			  			vr.cell = vr.rw.createCell(j);
			  			vr.cell.setCellValue("null");	
			  			
		  				break;
		  			}
		  			
		  			if(ME.equals("NSE") || ME.equals("BSE"))
		  			{
			  			String space = " ";
			  			vr.combine = vr.scrip.toUpperCase().concat(space).concat(IT).concat(space).concat(ME);
			  			System.out.println("combine "+vr.combine);
		  			}
		  			else
		  			{
		  				String space = " ";
		  				vr.combine = vr.scrip.toUpperCase().concat(space).concat(IT).concat(space).concat(ME).concat(space).concat(ManageWatchlist.getexactexpdate);
			  			System.out.println("combine "+vr.combine);
		  			}
		  			
		  			j=j+1;
		  			if(ME.equals("NSE-Derivatives"))
		  			{
		  				ME="NSE";
		  				String space = " ";
		  				vr.combine = vr.scrip.toUpperCase().concat(space).concat(IT).concat(space).concat(ME).concat(space).concat(ManageWatchlist.getexactexpdate);

		  			}
		  			if(ME.equals("BSE-Derivatives"))
		  			{
		  				ME="BSE";
		  				String space = " ";
		  				vr.combine = vr.scrip.toUpperCase().concat(space).concat(IT).concat(space).concat(ME).concat(space).concat(ManageWatchlist.getexactexpdate);

		  			}
		  					  			
		  			vr.rw = s.getRow(i);
		  			vr.cell = vr.rw.createCell(j);
		  			vr.cell.setCellValue(vr.combine);
		  					  			
		  			j=j+1;		  			  			
		  			vr.rw = s.getRow(i);
		  			vr.cell = vr.rw.createCell(j);
		  			vr.cell.setCellValue(functions.b);		  			  			
		  			
		  			log.log(LogStatus.INFO, "Scrip "+vr.scrip.toUpperCase()+ " got added.");
		  			break;
		  		}
	  			else
	  			{
	  				break;
	  			}
	  		}
	  	}
		FileOutputStream Of = new FileOutputStream(vr.srcfile);

	  	wb.write(Of);
	  	Of.close();
	}
	
	public void capturetotal() throws Exception
	{		
		///click funds >>equity/derivatives
		Actions act1 = new Actions(w);
		act1.moveToElement(w.findElement(By.linkText("Funds"))).click().perform();	
		WebElement Sublink1 = w.findElement(By.linkText("Equity/Derivatives"));
		Sublink1.click();
					
		//capture utilisation total
		w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).click();
		
////////////////////////////////
		passtakesnapshot();
		
		String checkmsg = "No Record Found";
				
		vr.scrip1 = s.getRow(i).getCell(j).getStringCellValue();
		System.out.println("scrip1 "+vr.scrip1);
		
		if(vr.scrip1 !=null &&  !vr.scrip1.isEmpty())
		{
			if(!w.getPageSource().contains(checkmsg))
			{
				if(w.getPageSource().contains(vr.scrip1))
				{
				System.out.println("in loop scrip1 " +vr.scrip1);
				String capture = w.findElement(By.xpath("//td[text()='"+vr.scrip1+"']/following-sibling::td[6]")).getText();
				vr.cap = Double.parseDouble(capture);
				System.out.println("first/cap " +vr.cap);
				}
				else
				{
					vr.cap = 0.0;
					System.out.println("cap of capture " +vr.cap);
				}
				vr.rw = s.getRow(i);
				vr.cell = vr.rw.createCell(variables.wputil);
				vr.cell.setCellValue(vr.cap);
			}
			else
			{
				vr.cap = 0.0;
				vr.rw = s.getRow(i);///////////////updated
				vr.cell = vr.rw.createCell(variables.wputil);///////////////updated
				vr.cell.setCellValue(vr.cap);///////////////updated
				System.out.println("Out of loop");
			}
		}
		else
		{
			
		}
				
	}
	
	@Test(priority=0)
	public void placeorder() throws Exception
	{	
  		FileInputStream f = new FileInputStream(srcfile);
		wb = new HSSFWorkbook(f);
		
		vr.s = wb.getSheet("M2M");
		
		///report creation
				report1 = new ExtentReports("D:\\Saleema\\utilisation.html");
				log = report1.startTest("Manage Watchlist");
			  	
				////call login function
			  	Login();
			  	
			  	///call manage watchlist
			  	manage();
			  		  	
			  	///call add scrip
			  	addloopscrips();
			  	
			  	///get multiple value from excel for utilization caln
		
			  	for(i=5;i<rc+1;i++)
			  	{
			  		Row row = s.getRow(i);
			  		
			  		String flag = s.getRow(i).getCell(0).getStringCellValue();
						
						if(flag.equals("Y"))
						{
					  		forreportscripname=s.getRow(i).getCell(1).getStringCellValue();
						  	System.out.println("reportscripname "+forreportscripname);
						  	log = report1.startTest("Place " +forreportscripname+ " Order");
						  	
			  				for(j=5;j<row.getLastCellNum();)
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
					  
					  for(int k=1;k<=rc1.size();k++)
					  {
						  vr.getscrip = wltable.findElement(By.name("Row " + (k) + ", Column 1")).getText();
						  System.out.println("get scrip " +vr.getscrip);
						  wltable.findElement(By.name("Row " + (k)+ ", Column 1")).click();
						  getxchnge = wltable.findElement(By.name("Row " + (k)+ ", Column 2")).getText();
						  System.out.println("get exchnage "+getxchnge);
						  vr.getprice = wltable.findElement(By.name("Row " + (k)+ ", Column 9")).getText();
						  System.out.println("get LTP "+vr.getprice);
						  
						  double chngegetprice = Double.parseDouble(vr.getprice); 
						  
						  if(vr.scrip.equals(vr.getscrip) && ME.equals(getxchnge))
						  {			
							  ///select action
							  j=j+1;
							  vr.action = vr.s.getRow(i).getCell(j).getStringCellValue();

							  if(vr.action.equals("BUY"))							  
							  {
								  robot.keyPress(KeyEvent.VK_F1);
							  }
							  else
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
							  double chngegetamt = Double.parseDouble(vr.getamt);
							  
							  if(chngegetamt==0)
							  {
								  vr.getamt = vr.getprice;								  						  
							  }
							  
							  ///get price perc
							  j=j+1;
							  vr.perc = (float) vr.s.getRow(i).getCell(j).getNumericCellValue();
							  if(vr.perc!=0)
							  {
								  vr.getamt1 = (Double.parseDouble(vr.getamt)*vr.perc)/100;
								  double chngestrnggetamt = Double.parseDouble(vr.getamt);
								  if(vr.action.equals("BUY"))
									{
										vr.total = chngestrnggetamt-vr.getamt1;
									}
									else
									{
										vr.total = chngestrnggetamt+vr.getamt1;
									}
								  vr.roundoff = (double) (Math.round(vr.total));
							  }
							  else
							  {
								  vr.roundoff = 0.0;
							  }
							  
							  log.log(LogStatus.INFO, "Actual price is " +vr.getamt);
							  log.log(LogStatus.INFO, "Modified price is " +vr.roundoff);	
							  								  								  
							  robot.keyPress(KeyEvent.VK_DELETE);	
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
							  							  
							  	  String msg2 = "Success. ABWS Order number is:";
								  for(int m=0;m<4;)
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
										  j=j+1;
										  if(getordno[0]!=null)
										  {
											  vr.rw = vr.s.getRow(i);
											  vr.cell = vr.rw.createCell(j);
											  vr.cell.setCellValue(getordno[0]);
										  }
										  else
										  {
											  vr.rw = vr.s.getRow(i);
											  vr.cell = vr.rw.createCell(j);
											  vr.cell.setCellValue("null");
										  }
										  										  
										  break;
									  }
									  else
									  {
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
	
	public void chktrad() throws Exception
	{
		robot.keyPress(KeyEvent.VK_ALT);
		Thread.sleep(500);
		robot.keyPress(KeyEvent.VK_R);
		Thread.sleep(500);
		robot.keyPress(KeyEvent.VK_O);
		Thread.sleep(1000);
		robot.keyRelease(KeyEvent.VK_ALT);
		
		log = report1.startTest("Order Status");
		log.log(LogStatus.INFO, "Open Order Report ");		
		
	  	FileOutputStream Of = new FileOutputStream(srcfile);
	  		  					
		for(i=5;i<xcelrc+1;)
	  	{	
			/////open report
			
			Row row = vr.s.getRow(i);			  
			String flag = vr.s.getRow(i).getCell(0).getStringCellValue();
			if(flag.equals("Y"))
			  {
				  for(j=1;j<row.getLastCellNum();)
				  {
					  	vr.getscrip = vr.s.getRow(i).getCell(j).getStringCellValue();
					  	System.out.println("get scrip " +vr.getscrip);
					  	
					  	j=j+1;
					  	getxchnge = vr.s.getRow(i).getCell(j).getStringCellValue();
					  	System.out.println("get exchnage "+getxchnge);
					  	
					  	j=j+1;
					  	getaction = vr.s.getRow(i).getCell(j).getStringCellValue();
					  	System.out.println("get action "+getaction);
					  	
					  	j=j+1;
					  	stclientcode = vr.s.getRow(i).getCell(j).getStringCellValue();
					  	
					  	j=j+1;
					  	vr.qty = (int) vr.s.getRow(i).getCell(j).getNumericCellValue();
					  	
					  	j=j+6;
					  	xcelordno = vr.s.getRow(i).getCell(j).getStringCellValue();
					  	System.out.println("xcel order no "+xcelordno);
					  	
					  	//enter client code
						WebElement rpttable1 = w.findElement(By.name("Order Report [All]"));
						rpttable1.findElement(By.id("1018")).sendKeys(stclientcode);
						robot.keyPress(KeyEvent.VK_F5);
						log.log(LogStatus.INFO, "Enter Client Code : "+stclientcode);
						
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
							  
							  if(capordno.equals(xcelordno))
							  {
								  if(getstat.equals("TRAD"))
								  {
									  price = rpttable.findElement(By.name("Row " + (m)+ ", Column 8")).getText();
									  System.out.println("mod price "+price);
									  
									  log.log(LogStatus.INFO, "Select Order no : "+capordno);
									  log.log(LogStatus.INFO, "Selected order no status is : "+getstat);
									  log.log(LogStatus.INFO, "Order price : "+price);
									  j=j+1;
									  vr.rw = vr.s.getRow(i);
									  vr.cell = vr.rw.createCell(j);
									  vr.cell.setCellValue(price);
									  calc();
									  close();
									  	robot.keyPress(KeyEvent.VK_ALT);
										Thread.sleep(500);
										robot.keyPress(KeyEvent.VK_R);
										Thread.sleep(500);
										robot.keyPress(KeyEvent.VK_O);
										Thread.sleep(1000);
										robot.keyRelease(KeyEvent.VK_ALT);
									  break;
								  }
								  else
								  {
									  vr.rw = vr.s.getRow(i);
									  vr.cell = vr.rw.createCell(vr.m2mstatus);
									  vr.cell.setCellValue("Status is " +getstat);
									  break;
								  }
								  
							  }
							  else
							  {
								  System.out.println("else...");
							  }
						  }							
						break;
				  }		
			  }		
			else
			{
				System.out.println("flag is N");
			}
			i=i+3;
			close();
			robot.keyPress(KeyEvent.VK_ALT);
			Thread.sleep(500);
			robot.keyPress(KeyEvent.VK_R);
			Thread.sleep(500);
			robot.keyPress(KeyEvent.VK_O);
			Thread.sleep(1000);
			robot.keyRelease(KeyEvent.VK_ALT);
	  	}	
		
		close();
		wb.write(Of);
		Of.close();		
	}
	
	public void calc() throws Exception
	{
		log = report1.startTest("Calculate M2M");
		robot.keyPress(KeyEvent.VK_ALT);
		Thread.sleep(500);
		robot.keyPress(KeyEvent.VK_R);
		Thread.sleep(500);
		robot.keyPress(KeyEvent.VK_P);
		Thread.sleep(1000);
		robot.keyRelease(KeyEvent.VK_ALT);
		log.log(LogStatus.INFO, "Open Positions Report ");
		
		//enter client code
		WebElement optable = w.findElement(By.name("Net Position"));
		optable.findElement(By.id("1018")).sendKeys(stclientcode);
		robot.keyPress(KeyEvent.VK_F5);
		
		//count no of rows
		WebElement nptable = w.findElement(By.name("Net Position")).findElement(By.id("37033"));			  
		List < WebElement > rc = nptable.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
		System.out.println("row count "+ rc.size());
		  
		for(int k=1;k<rc.size();k++)
		{				  
			 WebElement abc = nptable.findElement(By.name("Row " + (k) + ""));
			 capscrip = abc.findElement(By.name("Row " + (k) + ", Column 1")).getText();
			 System.out.println("cap util "+capscrip);
			 log.log(LogStatus.INFO, "Selected scrip is : "+capscrip);
			 
			 //capture ME for BSE
			 capME = abc.findElement(By.name("Row " + (k) + ", Column 2")).getText();
			 System.out.println("cap ME "+capME);
			 
			 if(capscrip.equals(vr.getscrip) && capME.equals(getxchnge))
			 {			
				 capLTP = abc.findElement(By.name("Row " + (k) + ", Column 3")).getText();
				 System.out.println("traded cap LTP "+capLTP);
				 log.log(LogStatus.INFO, "LTP is "+capLTP);
				 j=j+1;
				 vr.rw = vr.s.getRow(i);
				 vr.cell = vr.rw.createCell(j);
				 vr.cell.setCellValue(capLTP);			 
				 
				 capqty = abc.findElement(By.name("Row " + (k) + ", Column 8")).getText();
				 //System.out.println("cap qty "+capqty);			 
				 
				 capnettrdval = abc.findElement(By.name("Row " + (k) + ", Column 9")).getText();
				 //System.out.println("cap net traded value "+capnettrdval);
				 j=j+1;
				 vr.rw = vr.s.getRow(i);
				 vr.cell = vr.rw.createCell(j);
				 vr.cell.setCellValue(vr.qty);
				 
				 capunrealpl = abc.findElement(By.name("Row " + (k) + ", Column 13")).getText();
				 System.out.println("cap unrealised P/L "+capunrealpl);
				 log.log(LogStatus.INFO, "Unrealised P/L : "+capunrealpl);
				 double chngeunreal = Double.parseDouble(capunrealpl);
				 j=j+1;
				 vr.rw = vr.s.getRow(i);
				 vr.cell = vr.rw.createCell(j);
				 vr.cell.setCellValue(capunrealpl);		
				 
				 ///compare unrealised with actual calculation
				 double chngeprevLTP = Double.parseDouble(price);				 
				 double chngecapLTP = Double.parseDouble(capLTP);
///////////////////////////////////////////////////////////////////////////////////////////
				/*
				if(chngeprevLTP-chngecapLTP>0)
				{
					System.out.println("buy profit");
				}
				else
				{
					System.out.println("buy loss");
				}
				if(chngecapLTP-chngeprevLTP<0)
				{
					System.out.println("sell profit");
				}
				else
				{
					System.out.println("sell loss");
				}
				*/
////////////////////////////////////////////////////////////////////////////////////////////
				 if(getaction.equals("BUY"))
				 {
					 if(getxchnge.equals("NSE-FX")) ////for currency
					 {
						 if(chngecapLTP-chngeprevLTP>=0)	//////prevltp is order price and capltp is current ltp
						 {
							 calc=Math.round(((chngecapLTP-chngeprevLTP)*vr.qty*1000)*100.0)/100.0;
							 System.out.println("calculated unrealised buy profit "+calc);
							 log.log(LogStatus.INFO, "Buy Profit");
						 }
						 else
						 {
							 calc=Math.round(((chngecapLTP-chngeprevLTP)*vr.qty*1000)*100.0)/100.0;
							 System.out.println("calculated unrealised buy loss "+calc);
							 log.log(LogStatus.INFO, "Sell Loss");
						 }	
					 }
					 else ////for cash and fno
					 {
						 if(chngecapLTP-chngeprevLTP>=0)	//////prevltp is order price and capltp is current ltp
						 {
							 calc=Math.round((chngecapLTP-chngeprevLTP)*vr.qty*100.0)/100.0;
							 System.out.println("calculated unrealised buy profit "+calc);
							 log.log(LogStatus.INFO, "Buy Profit");
						 }
						 else
						 {
							 calc=Math.round((chngecapLTP-chngeprevLTP)*vr.qty*100.0)/100.0;
							 System.out.println("calculated unrealised buy loss "+calc);
							 log.log(LogStatus.INFO, "Sell Loss");
						 }
					 }
				 }
				 else if(getaction.equals("SELL"))
				 {
					 if(getxchnge.equals("NSE-FX")) ///for currency
					 {
						 if(chngeprevLTP-chngecapLTP>=0)
						 {
							 calc=Math.round(((chngeprevLTP-chngecapLTP)*vr.qty*1000)*100.0)/100.0;
							 System.out.println("calculated unrealised sell profit "+calc);
							 log.log(LogStatus.INFO, "Sell Profit");
						 }
						 else
						 {
							 calc=Math.round(((chngeprevLTP-chngecapLTP)*vr.qty*1000)*100.0)/100.0;
							 System.out.println("calculated unrealised sell loss "+calc);
							 log.log(LogStatus.INFO, "Buy Loss");
						 }	
					 }
					 else   ////for cash and fno
					 {
						 if(chngeprevLTP-chngecapLTP>=0)
						 {
							 calc=Math.round((chngeprevLTP-chngecapLTP)*vr.qty*100.0)/100.0;
							 System.out.println("calculated unrealised sell profit "+calc);
							 log.log(LogStatus.INFO, "Sell Profit");
						 }
						 else
						 {
							 calc=Math.round((chngeprevLTP-chngecapLTP)*vr.qty*100.0)/100.0;
							 System.out.println("calculated unrealised sell loss "+calc);
							 log.log(LogStatus.INFO, "Buy Loss");
						 }
					 }
				 }
				 j=j+1;
				 vr.rw = vr.s.getRow(i);
				 vr.cell = vr.rw.createCell(j);
				 vr.cell.setCellValue(calc);
				 
				 if(chngeunreal==calc)
				 {
					 System.out.println("unrealised correct");
					 log.log(LogStatus.INFO, "Calculated and Actual Unrealised P/L is equal");
					 
					 vr.rw = vr.s.getRow(i);
					 vr.cell = vr.rw.createCell(vr.m2mstatus);
					 vr.cell.setCellValue("Calc and actual Unrealised P/L is Equal");
				 }
				 else
				 {
					 System.out.println("unrealised not correct");
					 log.log(LogStatus.INFO, "Calculated and Actual Unrealised P/L is unequal");
					 
					 vr.rw = vr.s.getRow(i);
					 vr.cell = vr.rw.createCell(vr.m2mstatus);
					 vr.cell.setCellValue("Calc and actual Unrealised P/L is not Equal");
				 }
			 }
			 else
			 {
				 System.out.println("No scrip");
			 }
		}
		close();
	}
	
	@Test
	public void M2Mcalc() throws IOException, Exception
	{
		  ///report creation
		  report1 = new ExtentReports("D:\\testfile\\m2m.html");
		  log = report1.startTest("Login Details");
		  
		  ////call login function
		  Login();
		  
		  robot = new Robot();
		  
		  ////open watchlist
		  openWL();
		  //place order function
		  placeorder();
		  chktrad();
		  
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
	
	
	@Test(priority=4003)
	public void nouse()
	  {
		  report1.endTest(log);
		  report1.flush();
		  //w.get("D:\\testfile\\winapputil.html");				  
		  //w.get(srcfile);
	  }
}
