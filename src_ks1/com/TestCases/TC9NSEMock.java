package com.TestCases;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.AfterTest;
import org.testng.annotations.Test;

import com.library.functions;
import com.library.variables;
import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.LogStatus;

import Main.BSOrderXL;
import Utilisation.ManageWatchlist;

public class TC9NSEMock extends BSOrderXL
{
	public static HSSFSheet s;
	public static String popup1;
	public static String value;
	public static int i,j,k,rc;
	public static WebElement ele;
	public static String getexactexpdate;
	
	//Utilisation.ManageWatchlist cm = new Utilisation.ManageWatchlist();
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
				        //fnodate = s.getRow(8).getCell(2).getStringCellValue();
				        //currdate = s.getRow(8).getCell(3).getStringCellValue();
		  				//fn.getdate();	  				
		  				fn.getdate1();		  				
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
		
	@Test(priority=0)
	public void PlaceOrder() throws Exception 
	  {					
		FileInputStream f = new FileInputStream(srcfile);
		wb = new HSSFWorkbook(f);
		
		s = wb.getSheet("NSEMock");
						
		///report creation
		report1 = new ExtentReports("D:\\Saleema\\utilisation.html");
		log = report1.startTest("Manage Watchlist");
	  	
		////call login function
	  	Login();
	  	
	  	///call manage watchlist
	  	manage();
	  		  	
	  	///call add scrip
	  	addloopscrips();
	  	
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
	  					deltakesnapshot();
	  					
	  					///click watchlist>>place order
						Actions act = new Actions(w);
						act.moveToElement(w.findElement(By.linkText("Watchlist"))).click().perform();	
						WebElement Sublink = w.findElement(By.linkText("My Watchlists"));
						Sublink.click();
						log.log(LogStatus.INFO, "Click on Watchlist >> My Watchlists");
						
						String arrow = ">>";
						WebElement mytable = w.findElement(By.xpath("//*[@id='tabbed_area']/td/ul"));
						//To locate rows of table. 
						List < WebElement > rows_table = mytable.findElements(By.tagName("li"));
						//To calculate no of rows In table.
						int rows_count = rows_table.size();
						System.out.println("row count "+rows_count);
				    	if(rows_count>1)
				    	{
							//Loop will execute till the last row of table.
							for (int row1 = 0; row1 < rows_count;row1++) 
							{
								// To retrieve text from that specific cell.
								String scrip = rows_table.get(row1).getText();
								System.out.println("Cell Value of row number " + row1 +  " Is " + scrip);
								if(scrip.equals(">>"))
								{
									w.findElement(By.xpath("//div[@id='container']//a[contains(text(), '"+arrow+"')]")).click();
									break;
								}						    	     
							}
							w.findElement(By.xpath("//div[@id='container']//a[contains(text(), '"+vr.watchlistname+"')]")).click();
				    	}
				    	///select scrip
				    	j=j+1;
						String xcelscrip = s.getRow(i).getCell(j).getStringCellValue();
						System.out.println("data1 " +xcelscrip);
						w.findElement(By.xpath(".//input[starts-with(@value,'"+xcelscrip+"')]")).click();
						
						j=j+1;
						String getexccode = w.findElement(By.xpath(".//td[input[starts-with(@value,'"+xcelscrip+"')]]/following-sibling::td[3]")).getText();
						System.out.println("getexchange code "+getexccode);
						
						vr.rw = s.getRow(i);
						vr.cell = vr.rw.createCell(j);
						vr.cell.setCellValue(getexccode);						
						
						////click on place order
						w.findElement(By.partialLinkText("PLACE ORDER")).click();
						log.log(LogStatus.INFO, "Click on Place Order");
						
						///get scripname
						vr.scripname = w.findElement(By.id("stk_name")).getAttribute("value");
						log.log(LogStatus.INFO, "Selected Scrip is " +vr.scripname);
						
						///select buy/sell
						j=j+1;
						value = s.getRow(i).getCell(j).getStringCellValue();
						System.out.println("boolean " +value);
						w.findElement(By.xpath(".//input[@value='"+value+"']")).click();
						log.log(LogStatus.INFO, "Selected Value is " +value);
											
						////enter quantity
						j=j+1;
						vr.quantity = (int) s.getRow(i).getCell(j).getNumericCellValue();  	
						System.out.println("data1 " +vr.quantity);
						
						if(w.findElement(By.id("stk_qty")).isEnabled())
						{
							////chk for cash
							w.findElement(By.id("stk_qty")).sendKeys(String.valueOf(vr.quantity));
							log.log(LogStatus.INFO, "Entered Quantity is: " +vr.quantity);
							
							///multiple
							j=j+1;
							vr.multiple = (float) s.getRow(i).getCell(j).getNumericCellValue();  	
							System.out.println("multiple " +vr.multiple);
							
							///required margin calculation
							j=j+1;
							int perc = (int) s.getRow(i).getCell(j).getNumericCellValue();
							vr.pr1 = w.findElement(By.id("stk_price")).getAttribute("value");
							w.findElement(By.id("stk_price")).clear();
							vr.price = (Double.parseDouble(vr.pr1)*perc)/100;
							double chngestrngpr = Double.parseDouble(vr.pr1);
							if(value.equals("BUY"))
							{
								vr.total = chngestrngpr-vr.price;
							}
							else
							{
								vr.total = chngestrngpr+vr.price;
							}
							vr.roundoffpr = (double) (Math.round(vr.total));
							w.findElement(By.id("stk_price")).sendKeys(String.valueOf(vr.roundoffpr));
							log.log(LogStatus.INFO, "Actual price is " +vr.pr1);
							log.log(LogStatus.INFO, "Modified price is " +vr.roundoffpr);
						}
						else
						{
							////chk for fno and currency
							w.findElement(By.id("stk_lot")).sendKeys(String.valueOf(vr.quantity));
							log.log(LogStatus.INFO, "Entered No of lots is: " +vr.quantity);
							
							///multiple
							j=j+1;
							vr.multiple = (float) s.getRow(i).getCell(j).getNumericCellValue();  	
							System.out.println("multiple  " +vr.multiple);
							
							///required margin calculation
							j=j+1;
							int perc = (int) s.getRow(i).getCell(j).getNumericCellValue();
							vr.pr1 = w.findElement(By.id("stk_price")).getAttribute("value");
							w.findElement(By.id("stk_price")).clear();
							vr.price = (Double.parseDouble(vr.pr1)*perc)/100;
							double chngestrngpr = Double.parseDouble(vr.pr1);
							if(value.equals("BUY"))
							{
								vr.total = chngestrngpr-vr.price;
							}
							else
							{
								vr.total = chngestrngpr+vr.price;
							}
							vr.roundoffpr = (double) (Math.round(vr.total));
							w.findElement(By.id("stk_price")).sendKeys(String.valueOf(vr.roundoffpr));
							log.log(LogStatus.INFO, "Actual price is " +vr.pr1);
							log.log(LogStatus.INFO, "Modified price is " +vr.roundoffpr);
							
							///write actual price value to excel
							j=j+1;
							vr.rw = s.getRow(i);
							vr.cell = vr.rw.createCell(j);
							vr.cell.setCellValue(vr.pr1);
							
							///write modified price value to excel
							j=j+1;
							vr.rw = s.getRow(i);
							vr.cell = vr.rw.createCell(j);
							vr.cell.setCellValue(vr.roundoffpr);
							
							///select order type
							j=j+1;
							String ordmrk = s.getRow(i).getCell(j).getStringCellValue();
							Select OT = new Select(w.findElement(By.id("del")));
							OT.selectByVisibleText(ordmrk);
							
						}
	  				} ///j loop
				} //flag
	  	} //i loop
				
	  }

	@AfterTest
	public void aftertest()
	{
		//w.close();
	}
		
	@Test(priority=4001)
	public void nouse()
	  {
		  report1.endTest(log);
		  report1.flush();
		  w.get("D:\\Saleema\\NSEMock.html");
		  w.get(srcfile);
	  }
}
