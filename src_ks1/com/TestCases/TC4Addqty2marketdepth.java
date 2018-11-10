package com.TestCases;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.TimeZone;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.Test;

import com.library.functions;
import com.library.variables;
import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import Main.BSOrderXL;
import Utilisation.ManageWatchlist;

public class TC4Addqty2marketdepth
{
	  public static WebDriver w;
	  public static ExtentReports report1;
	  public static ExtentTest log;
	  String URL,Uname,Pwd;
	  int AccCode,i,j,rc;
	  public static WebElement ele;
	  public static String value;
	  
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
			vr.watchlistname = vr.s.getRow(1).getCell(5).getStringCellValue();;  	
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
			//FileOutputStream Of = new FileOutputStream(srcfile);
						
			rc = vr.s.getLastRowNum() - vr.s.getFirstRowNum();
		  	System.out.println("rc "+rc);
		  	  	
		  	for(i=5;i<rc+1;i++)
		  	{
		  		Row row = vr.s.getRow(i);
		  		for(j=0;j<row.getLastCellNum();)
		  		{
		  			String flag = vr.s.getRow(i).getCell(j).getStringCellValue();
		  			
		  			if(flag.equals("Y"))
		  			{
			  			w.findElement(By.partialLinkText("ADD")).click();
			  			
			  			j=j+2;
			  			vr.scrip = vr.s.getRow(i).getCell(j).getStringCellValue();
			  			w.findElement(By.id("search_list")).sendKeys(vr.scrip);
			  			
			  			j=j+1;
			  			String ME = vr.s.getRow(i).getCell(j).getStringCellValue();
			  			w.findElement(By.id("me")).sendKeys(ME);
			  				  				  			
			  			j=j+1;
			  			String IT = vr.s.getRow(i).getCell(j).getStringCellValue();
			  			w.findElement(By.id("inst_type")).sendKeys(IT);
			  				  			
			  			w.findElement(By.partialLinkText("SUBMIT")).click();	  								  						  								
		
			  			//to check if scrip exist
			  			if(w.getPageSource().contains(vr.scrip.toUpperCase()))
			        	{	  				
			  				getdate();
			  				
			        	}
			  			else
			  			{
			  				w.findElement(By.partialLinkText("CANCEL")).click();
				  			log.log(LogStatus.INFO, "No matching stocks found for the Scrip " +vr.scrip.toUpperCase());
				  			
				  			j=j+1;
				  			vr.rw = vr.s.getRow(i);
				  			vr.cell = vr.rw.createCell(j);
				  			vr.cell.setCellValue("null");
				  			
				  			j=j+1;		  			  			
				  			vr.rw = vr.s.getRow(i);
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
				  			vr.combine = vr.scrip.toUpperCase().concat(space).concat(IT).concat(space).concat(ME).concat(space).concat(vr.getexactexpdate);
				  			System.out.println("combine "+vr.combine);
			  			}
			  			
			  			j=j+1;
			  			if(ME.equals("NSE-Derivatives"))
			  			{
			  				ME="NSE";
			  				String space = " ";
				  			vr.combine = vr.scrip.toUpperCase().concat(space).concat(IT).concat(space).concat(ME).concat(space).concat(vr.getexactexpdate);

			  			}
			  			if(ME.equals("BSE-Derivatives"))
			  			{
			  				ME="BSE";
			  				String space = " ";
				  			vr.combine = vr.scrip.toUpperCase().concat(space).concat(IT).concat(space).concat(ME).concat(space).concat(vr.getexactexpdate);

			  			}
			  					  			
			  			vr.rw = vr.s.getRow(i);
			  			vr.cell = vr.rw.createCell(j);
			  			vr.cell.setCellValue(vr.combine);
			  					  			
			  			j=j+1;		  			  			
			  			vr.rw = vr.s.getRow(i);
			  			vr.cell = vr.rw.createCell(j);
			  			vr.cell.setCellValue(vr.b);		  			  			
			  			
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
		  	vr.wb.write(Of);
		  	Of.close();
		}
	  
	  public void getdate()
		{
			List<WebElement> myList= w.findElements(By.xpath(".//td[input[@type = 'checkbox']]"));
			List<WebElement> myListele = w.findElements(By.xpath("//input[@name='stockcode' and @type = 'checkbox']"));
			System.out.println("Total elements :" +myList.size());
			
	        Date currDate = Calendar.getInstance().getTime();
		   	SimpleDateFormat dateFormat = new SimpleDateFormat("ddMMMyy");
			SimpleDateFormat dateFormatformonth = new SimpleDateFormat("MMM");
			SimpleDateFormat dateFormatforyear = new SimpleDateFormat("yy");
			
			String getcurrdate = dateFormat.format(currDate);
			System.out.println("curr date "+getcurrdate);
		    	 
			Calendar calendar = Calendar.getInstance(TimeZone.getDefault());
			int year = calendar.get(Calendar.YEAR);
			int date = calendar.get(Calendar.DAY_OF_MONTH);
			String month = dateFormatformonth.format(currDate);
			String yr = dateFormatforyear.format(currDate);
			vr.dte = String.valueOf(date);
			
	        String arr[]=new String[myList.size()];

	 	    for(int i=0;i<myList.size();i++)
		 	{			
	 	    	WebElement ele;
				arr[i]= myList.get(i).getText();		
				String actualscrip = arr[i];
	 	    	ele = myListele.get(i);
	 	    	ele.click();
				
				System.out.println("actualscrip "+actualscrip);
				String[] getscripexpirydate = actualscrip.split("]" ,0);
				String gsed = getscripexpirydate[1];
				System.out.println("gsed "+gsed);
				if(gsed.startsWith("["))
				{
					String[]removebrac = gsed.split("\\[");
					System.out.println("1 "+removebrac[1]);
					vr.getexactexpdate = removebrac[1];
					        		
					String[] getonlydate = vr.getexactexpdate.split("(?<=[A-Z])(?=[0-9])|(?<=[0-9])(?=[A-Z])");
					vr.cmprdate = getonlydate[0];
					vr.cmprmonth = getonlydate[1];
					vr.cmpryr = getonlydate[2];
					
					//int da = Integer.parseInt(getonlydate[0]);
					int result = vr.dte.compareTo(vr.cmprdate);
					System.out.println("result "+result);
					//if(da<=date)
					if(result<=0 || result>0)
					//if(removebrac[1].equals("26JUL18") || removebrac[1].equals("27JUL18"))
					{
						if(month.toUpperCase().equals(vr.cmprmonth)) //comment
	        			{
							if(yr.equals(vr.cmpryr))
				 			{
								vr.b = ele.getAttribute("value");
					  			System.out.println("getvalue " +vr.b);
					  			w.findElement(By.partialLinkText("SUBMIT")).click();
					  			break;
		        			}
	        			}
						else //comment
						{
							ele.click(); //comment
						}		
					}
					else
					{
						ele.click();
					}
				}
				else
				{				
					vr.b = ele.getAttribute("value");
		  			System.out.println("getvalue " +vr.b);
		  			w.findElement(By.partialLinkText("SUBMIT")).click();
					break;
				}
		 	}
		}
	  
	  @BeforeSuite
		public void beforeSuite() throws Exception
		{	  	
			  	System.setProperty("webdriver.chrome.driver", "D:\\Selenium software\\chromedriver.exe");
				w = new ChromeDriver();
				
				FileInputStream f = new FileInputStream(vr.srcfile);
				vr.wb = new HSSFWorkbook(f);
				vr.s = vr.wb.getSheet("Config");
				
				/// get URL
				URL = vr.s.getRow(2).getCell(1).getStringCellValue();
				w.get(URL);
				w.manage().window().maximize();
				w.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
				
				///report creation
				report1 = new ExtentReports("D:\\Saleema\\placemktdepth.html");
				log = report1.startTest("Manage Watchlist");
				
				Uname = vr.s.getRow(4).getCell(2).getStringCellValue();
			  	w.findElement(By.id("userid")).sendKeys(Uname);
			  	log.log(LogStatus.INFO, "Enter User name " +Uname);
			  	
			  	///enter password
			  	Pwd  = vr.s.getRow(4).getCell(3).getStringCellValue();
				w.findElement(By.id("pwd")).sendKeys(Pwd);
				log.log(LogStatus.INFO, "Enter Password ******");
				
				///enter access code
				AccCode= (int) vr.s.getRow(4).getCell(4).getNumericCellValue();
				w.findElement(By.id("otp")).sendKeys(String.valueOf(AccCode));
				log.log(LogStatus.INFO, "Enter Security Key/Access Code ****");
				
				///click on login
				//w.findElement(By.partialLinkText("Login")).click();
				w.findElement(By.xpath("//*[@id='tbdiv']/span")).click();
				//w.findElement(By.xpath("//*[@id='tbdiv']/span")).click();

				System.out.println("Logged in Successfully");
				log.log(LogStatus.INFO, "Logged in Successfully");
				
				//md.Placeorder();		
		}
	  
	  @Test(priority=0)
		public void addordtomktdepth() throws Exception
		{
			  FileInputStream f = new FileInputStream(vr.srcfile);
			  vr.wb = new HSSFWorkbook(f);
			  
			  vr.s = vr.wb.getSheet("Mktdepth");
			  
			  manage();
			  addloopscrips();
			  //md.Placeorder();
			  
			///report creation
				report1 = new ExtentReports("D:\\Saleema\\placemktdepth.html");
								
				rc = vr.s.getLastRowNum() - vr.s.getFirstRowNum();
			  	System.out.println("rc "+rc);
			  	
			  	///place order to add market depth
				
			  	for(i=5;i<rc+1;i++)
			  	{
			  		Row row = vr.s.getRow(i);
			  		 
			  		String flag = vr.s.getRow(i).getCell(0).getStringCellValue();
						
						if(flag.equals("Y"))
						{
							vr.forreportscripname=vr.s.getRow(i).getCell(1).getStringCellValue();
						  	System.out.println("reportscripname "+vr.forreportscripname);
						  	log = report1.startTest("Add quantity to market depth");
						  	
			  				for(j=5;j<row.getLastCellNum();)
			  				{	
			  			  		vr.scrip1 = vr.s.getRow(i).getCell(5).getStringCellValue();						  			  					
			  					
							  	///click watchlist>>place order
								Actions act = new Actions(w);
								act.moveToElement(w.findElement(By.linkText("Watchlist"))).click().perform();	
								WebElement Sublink = w.findElement(By.linkText("My Watchlists"));
								Sublink.click();
								log.log(LogStatus.INFO, "Click on Watchlist >> My Watchlists");
								
								///watchlistname									
								System.out.println("watchlistname "+vr.watchlistname);
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////								
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
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////								
									w.findElement(By.xpath("//div[@id='container']//a[contains(text(), '"+vr.watchlistname+"')]")).click();
						    	}
								///select scrip
								j=j+1;
								String xcelscrip = vr.s.getRow(i).getCell(j).getStringCellValue();
								if(xcelscrip.equals("null") || xcelscrip.equals(""))
								{
									log.log(LogStatus.INFO, "Scrip is not available");								
								}
								else
								{
								System.out.println("data1 " +xcelscrip);
								w.findElement(By.xpath(".//input[starts-with(@value,'"+xcelscrip+"')]")).click();
								
								j=j+1;
								String getexccode = w.findElement(By.xpath(".//td[input[starts-with(@value,'"+xcelscrip+"')]]/following-sibling::td[3]")).getText();
								System.out.println("getexchange code "+getexccode);
								
								vr.rw = vr.s.getRow(i);
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
								vr.action = vr.s.getRow(i).getCell(j).getStringCellValue();
								System.out.println("boolean " +vr.action);
								w.findElement(By.xpath(".//input[@value='"+vr.action+"']")).click();
								log.log(LogStatus.INFO, "Selected Value is " +vr.action);
													
								////enter quantity
								j=j+1;
								vr.quantity = (int) vr.s.getRow(i).getCell(j).getNumericCellValue();  	
								System.out.println("data1 " +vr.quantity);
								//wb.close();
								
								if(w.findElement(By.id("stk_qty")).isEnabled())
								{
									w.findElement(By.id("stk_qty")).sendKeys(String.valueOf(vr.quantity));
									log.log(LogStatus.INFO, "Entered Quantity is: " +vr.quantity);															
									
									///price calculation for adding buy and sell qty to market depth
									j=j+1;
									int perc = (int) vr.s.getRow(i).getCell(j).getNumericCellValue();
									vr.pr1 = w.findElement(By.id("stk_price")).getAttribute("value");
									w.findElement(By.id("stk_price")).clear();
									vr.price = (Double.parseDouble(vr.pr1)*perc)/100;
									double chngestrngpr = Double.parseDouble(vr.pr1);
									if(vr.action.equals("BUY"))
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
									w.findElement(By.id("stk_lot")).sendKeys(String.valueOf(vr.quantity));
									log.log(LogStatus.INFO, "Entered No of lots is: " +vr.quantity);
									
									///price calculation for adding buy and sell qty to market depth
									j=j+1;
									int perc = (int) vr.s.getRow(i).getCell(j).getNumericCellValue();
									vr.pr1 = w.findElement(By.id("stk_price")).getAttribute("value");
									w.findElement(By.id("stk_price")).clear();
									vr.price = (Double.parseDouble(vr.pr1)*perc)/100;
									double chngestrngpr = Double.parseDouble(vr.pr1);
									if(vr.action.equals("BUY"))
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
								
								///place order
								w.findElement(By.partialLinkText("PLACE ORDER")).click();
								log.log(LogStatus.INFO, "Click on Place Order button");
								
								///alert
								w.switchTo().alert().accept();
								log.log(LogStatus.INFO, "Click OK on alert");
														
								///check for any error
								String popup = "Error";
								if (w.getPageSource().contains(popup))
								{
									vr.popup1 = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
									System.err.println("Error displayed as " +vr.popup1);
									log.log(LogStatus.ERROR, "" +vr.popup1);
									if(w.getPageSource().contains("BACK"))
									{
										w.findElement(By.partialLinkText("BACK")).click();
										j=j+1;
										vr.rw = vr.s.getRow(i);
										vr.cell = vr.rw.createCell(j);								
										vr.cell.setCellValue("null");
										
									}
									else
									{
										System.out.println("...");
									}
									vr.rw = vr.s.getRow(i);
									vr.cell = vr.rw.createCell(vr.mdstatus);
									vr.cell.setCellValue(vr.popup1);
									
								}
								else
								{	
									Thread.sleep(1000);
									////capture order no
									vr.stckno = w.findElement(By.xpath("//*[@id='container']/div/p/span")).getText();
							 	    System.out.println("stckno...... " +vr.stckno);
							 	    log.log(LogStatus.INFO, "Order No is " +vr.stckno);
							 	    
									j=j+1;
									vr.rw = vr.s.getRow(i);
									vr.cell = vr.rw.createCell(j);
									vr.cell.setCellValue(vr.stckno);	
																										
										w.findElement(By.partialLinkText("Order Status")).click();
										w.findElement(By.xpath(".//input[starts-with(@value,'"+vr.stckno+"')]")).click();
										String status = w.findElement(By.xpath("//td[text()='"+vr.stckno+"']/following-sibling::td[6]")).getText();
										
										vr.rw = vr.s.getRow(i);
										vr.cell = vr.rw.createCell(vr.mdstatus);
										vr.cell.setCellValue(status);
										
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
			  	
			  	FileOutputStream Of = new FileOutputStream(vr.srcfile);
			  	vr.wb.write(Of);
				Of.close();
		}
	  
	  @AfterTest
		public void aftertest()
		{
			w.close();
		}
	  
	  @Test(priority=4004)
		public void nouse()
		  {
			  report1.endTest(log);
			  report1.flush();
			  w.get("D:\\Saleema\\placemktdepth.html");
			  w.get(vr.srcfile);
		  }
	  
}
