package Testcases;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.testng.TestListenerAdapter;
import org.testng.TestNG;
import org.testng.annotations.AfterTest;
import org.testng.annotations.Test;

import com.library.functions;
import com.library.variables;
import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.LogStatus;

import Main.BSOrderXL;
import Utilisation.ManageWatchlist;

public class TC2Comparemaxutil extends BSOrderXL
{
	public static int i,j,k,rc;
	public static String combinemng;
	public static String scripvalmng;
	public static HSSFSheet scu;
	public static WebElement ele;
	public static String getexactexpdate,actiontype;
	public static double buymaxamt = 0;
	public static double sellmaxamt=0;
	public static double bma;
	public static double sma;
	public static double res;
	
	//com.utilisation.ManageWatchlist cm = new com.utilisation.ManageWatchlist();
	com.library.functions fn = new com.library.functions();
	com.library.variables vr = new com.library.variables();

	public static void main(String[] args) 
	  {
		  TestListenerAdapter tla = new TestListenerAdapter();
		    TestNG testng = new TestNG();
		     Class[] classes = new Class[]
		     {
		    	 TC2Comparemaxutil.class     
		     };
		     testng.setTestClasses(classes);
		     testng.addListener(tla);
		     testng.run();		  	  
	  	}
	
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
		vr.watchlistname = scu.getRow(1).getCell(5).getStringCellValue();  	
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
		FileOutputStream Of = new FileOutputStream(srcfile);
		int i;
		int j;
		
		rc = scu.getLastRowNum() - scu.getFirstRowNum();
	  	System.out.println("rc "+rc);
	  	  	
	  	for(i=5;i<rc+1;i++)
	  	{
	  		Row row = scu.getRow(i);
	  		for(j=0;j<row.getLastCellNum();)
	  		{
	  			String flag = scu.getRow(i).getCell(j).getStringCellValue();
	  			
	  			if(flag.equals("Y"))
	  			{
		  			w.findElement(By.partialLinkText("ADD")).click();
		  			
		  			j=j+2;
		  			vr.scrip = scu.getRow(i).getCell(j).getStringCellValue();
		  			w.findElement(By.id("search_list")).sendKeys(vr.scrip);
		  			
		  			j=j+1;
		  			String ME = scu.getRow(i).getCell(j).getStringCellValue();
		  			w.findElement(By.id("me")).sendKeys(ME);
		  				  				  			
		  			j=j+1;
		  			String IT = scu.getRow(i).getCell(j).getStringCellValue();
		  			w.findElement(By.id("inst_type")).sendKeys(IT);
		  				  			
		  			w.findElement(By.partialLinkText("SUBMIT")).click();	  								  						  								
	
		  			//to check if scrip exist
		  			if(w.getPageSource().contains(vr.scrip.toUpperCase()))
		        	{	  				
		  				fn.getdate();
		  				
		        	}
		  			else
		  			{
		  				w.findElement(By.partialLinkText("CANCEL")).click();
			  			log.log(LogStatus.INFO, "No matching stocks found for the Scrip " +vr.scrip.toUpperCase());
			  			
			  			j=j+1;
			  			vr.rw = scu.getRow(i);
			  			vr.cell = vr.rw.createCell(j);
			  			vr.cell.setCellValue("null");
			  			
			  			j=j+1;		  			  			
			  			vr.rw = scu.getRow(i);
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
		  					  			
		  			vr.rw = scu.getRow(i);
		  			vr.cell = vr.rw.createCell(j);
		  			vr.cell.setCellValue(vr.combine);
		  					  			
		  			j=j+1;		  			  			
		  			vr.rw = scu.getRow(i);
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
		String checkmsg = "No Record Found";
				
		vr.scrip1 = scu.getRow(i).getCell(j).getStringCellValue();
		System.out.println("scrip1 "+vr.scrip1);
		
		if(vr.scrip1.equals("null") || vr.scrip1.equals("") )
		{
			System.out.println("Scrip is not available ");
		}
		else
		{
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
			}
			else
			{
				vr.cap = 0.0;
				System.out.println("Out of loop");
			}
		}
		else
		{
			
		}
	  }		
	}	
	
		
	@Test(priority=0)
	public void PlaceOrder() throws Exception 
	  {
		FileInputStream f = new FileInputStream(vr.srcfile);
		wb = new HSSFWorkbook(f);
	  	FileOutputStream Of = new FileOutputStream(srcfile);

		scu = wb.getSheet("Compareutil");
				
		///report creation
		report1 = new ExtentReports("D:\\testfile\\cmprmaxutil.html");
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
	  		Row row = scu.getRow(i);
	  		 
	  		String flag = scu.getRow(i).getCell(0).getStringCellValue();
				
				if(flag.equals("Y"))
				{
					vr.forreportscripname=scu.getRow(i).getCell(1).getStringCellValue();
				  	System.out.println("reportscripname "+vr.forreportscripname);
				  	actiontype = scu.getRow(i).getCell(8).getStringCellValue();
				  	log = report1.startTest("Place " +vr.forreportscripname+ " " +actiontype+ " Order");
	  		
	  				for(j=5;j<row.getLastCellNum();)
	  				{	
	  			  		vr.scrip1 = scu.getRow(i).getCell(5).getStringCellValue();
	  			  		
					  	///call capture total function to capture utilisation value
	  					capturetotal();
						
	  					vr.rw = scu.getRow(i);
	  					vr.cell = vr.rw.createCell(variables.cwputil);
	  					vr.cell.setCellValue(vr.cap);
	  					
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
						String xcelscrip = scu.getRow(i).getCell(j).getStringCellValue();
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
						
						vr.rw = scu.getRow(i);
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
						vr.action = scu.getRow(i).getCell(j).getStringCellValue();
						System.out.println("boolean " +vr.action);
						w.findElement(By.xpath(".//input[@value='"+vr.action+"']")).click();
						log.log(LogStatus.INFO, "Selected Value is " +vr.action);
											
						////enter quantity
						j=j+1;
						vr.quantity = (int) scu.getRow(i).getCell(j).getNumericCellValue();  	
						System.out.println("data1 " +vr.quantity);
						//wb.close();
						
						if(w.findElement(By.id("stk_qty")).isEnabled())
						{
							w.findElement(By.id("stk_qty")).sendKeys(String.valueOf(vr.quantity));
							log.log(LogStatus.INFO, "Entered Quantity is: " +vr.quantity);
							
							///multiple
							j=j+1;
							vr.multiple = (float) scu.getRow(i).getCell(j).getNumericCellValue();  	
							System.out.println("multiple " +vr.multiple);
							
							///required margin calculation
							j=j+1;
							int perc = (int) scu.getRow(i).getCell(j).getNumericCellValue();
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
							vr.RM = ((vr.quantity * vr.roundoffpr ) / vr.multiple);
							vr.roundoff = (Math.round(vr.RM*100.0)/100.0);
							log.log(LogStatus.INFO, "Required Margin is " +vr.roundoff);
							System.out.println("roundoff " +vr.roundoff);
						}
						else
						{
							w.findElement(By.id("stk_lot")).sendKeys(String.valueOf(vr.quantity));
							log.log(LogStatus.INFO, "Entered No of lots is: " +vr.quantity);
							
							///multiple
							j=j+1;
							vr.multiple = (float) scu.getRow(i).getCell(j).getNumericCellValue();  	
							System.out.println("multiple  " +vr.multiple);
							
							///required margin calculation
							j=j+1;
							int perc = (int) scu.getRow(i).getCell(j).getNumericCellValue();
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
							String qty = w.findElement(By.id("stk_qty")).getAttribute("value");
							log.log(LogStatus.INFO, "Quantity is: " +qty);
							vr.quantity = Integer.parseInt(qty);
							vr.RM = ((vr.quantity * vr.roundoffpr ) / vr.multiple);
							vr.roundoff = (Math.round(vr.RM*100.0)/100.0);
							log.log(LogStatus.INFO, "Required Margin is " +vr.roundoff);
							System.out.println("roundoff " +vr.roundoff);
						}
						
						///write actual price value to excel
						j=j+1;
						vr.rw = scu.getRow(i);
						vr.cell = vr.rw.createCell(j);
						vr.cell.setCellValue(vr.pr1);
						
						///write modified price value to excel
						j=j+1;
						vr.rw = scu.getRow(i);
						vr.cell = vr.rw.createCell(j);
						vr.cell.setCellValue(vr.roundoffpr);
																														
						///chk for AMO
						j=j+1;
						String amochk = scu.getRow(i).getCell(j).getStringCellValue();
						System.out.println("amochk "+amochk);
						if(amochk.equals("YES"))
						{
							w.findElement(By.id("chk_amo")).click();
							log.log(LogStatus.INFO, "Select Place AMO checkbox");
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
								vr.rw = scu.getRow(i);
								vr.cell = vr.rw.createCell(j);								
								vr.cell.setCellValue("null");
								
							}
							else
							{
								System.out.println("...");
							}
							vr.rw = scu.getRow(i);
							vr.cell = vr.rw.createCell(variables.cwstatus);
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
							vr.rw = scu.getRow(i);
							vr.cell = vr.rw.createCell(j);
							vr.cell.setCellValue(vr.stckno);
						
							///click funds >>equity/derivatives
							Actions act1 = new Actions(w);
							act1.moveToElement(w.findElement(By.linkText("Funds"))).click().perform();	
							WebElement Sublink1 = w.findElement(By.linkText("Equity/Derivatives"));
							Sublink1.click();
														
							//click margin utilised cash&FNO
							w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).click();
							
							if(w.getPageSource().contains(vr.scrip1))
							{
								///get util amount of specific scrip				
								vr.total1 = w.findElement(By.xpath("//td[text()='"+vr.scrip1+"']/following-sibling::td[6]")).getText();
								System.out.println("total1 " +vr.total1);
								vr.tot = Double.parseDouble(vr.total1);
								System.out.println("cap amt " +vr.tot);
								//Double match = tot - cap;
								//System.out.println("match " +match);
								//matchround = Math.round(match*100.0)/100.0;
								//System.out.println("f amt" +matchround);
																																
								vr.rw = scu.getRow(i);
								vr.cell = vr.rw.createCell(variables.cwpactutil);
								vr.cell.setCellValue(vr.roundoff);
								
								//Pcmprmaxutil();
									double result;
									vr.action = scu.getRow(i).getCell(8).getStringCellValue();
									if(vr.action.equals("BUY")) //&& w.getPageSource().contains(scrip1))
									{						
										bma = vr.roundoff;
										System.out.println("buy action total" +bma);
										if(vr.tot>=bma)
										{
											bma=vr.tot;
											System.out.println("buy action total" +vr.tot);
										}			
									}
									else if (vr.action.equals("SELL"))// && w.getPageSource().contains(scrip1))
									{
										sma = vr.roundoff;
										System.out.println("sell action total" +sma);
										if(vr.tot>=sma)
										{
											sma=vr.tot;
											System.out.println("sell action total" +vr.tot);
										}
									}
			
										if(bma>=sma)
										{
											result = bma;
											System.out.println("buyamt "+bma);
											bma=0;
										}
										else
										{
											result = sma;
											System.out.println("sellamt "+sma);
											sma=0;
										}	
									
									vr.rw = scu.getRow(i);
									vr.cell = vr.rw.createCell(variables.cwpmaxutil);
									vr.cell.setCellValue(result);
							}
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
	  	wb.write(Of);
		Of.close();
	  }
	
	@Test(priority=1)
	public void ChangeOrder() throws Exception 
	{
	  	FileOutputStream Of = new FileOutputStream(srcfile);
	  		  		 
	  		for(i=5;i<rc+1;i++)
		  	{	
	  		Row row = scu.getRow(i);
	  		String flag = scu.getRow(i).getCell(0).getStringCellValue();
				
				if(flag.equals("Y"))
				{
					vr.forreportscripname=scu.getRow(i).getCell(1).getStringCellValue();
				  	System.out.println("reportscripname "+vr.forreportscripname);
				  	log = report1.startTest("Change " +vr.forreportscripname+ " " +actiontype+ " Order");
 					  	
	  				for(j=13;j<row.getLastCellNum();)
	  				{	  					  						  						  					
	  			  		vr.scrip1 = scu.getRow(i).getCell(5).getStringCellValue();
	  			  		vr.multiple = (float) scu.getRow(i).getCell(10).getNumericCellValue();
						///to check for any error
						String amocheck1 = "Error";
					  	if(w.getPageSource().contains(amocheck1))
					  	{
					  		String capturemsg = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
					  		log.log(LogStatus.ERROR, "" +capturemsg);
					  		
					  	}
					  	
					  	else
					  	{
					  		
					  	///click watchlist>>place order
						Actions act1 = new Actions(w);
						act1.moveToElement(w.findElement(By.linkText("Reports"))).click().perform();	
						WebElement Sublink1 = w.findElement(By.linkText("Order Status"));
						Sublink1.click();
						log.log(LogStatus.INFO, "Click on Reports >> Order Status");
						
						vr.roundoffpr = scu.getRow(i).getCell(j).getNumericCellValue();
						//scripchk = scu.getRow(i).getCell(5).getStringCellValue();
						
						//check stkno
						j=j+2;
						vr.stckno = scu.getRow(i).getCell(j).getStringCellValue();
						System.out.println("manage stk no " +vr.stckno);
						if(vr.stckno.equals("null")|| vr.stckno.equals(""))
						{
							System.out.println("Order no not generated as unable to place order.");
							log.log(LogStatus.ERROR, "Order no not generated as unable to modify order.");							
						}
						else
						{
						if(vr.stckno !=null &&  !vr.stckno.isEmpty())
						{
							if(w.getPageSource().contains(vr.stckno))
							{
								w.findElement(By.xpath(".//input[starts-with(@value,'"+vr.stckno+"')]")).click();
								log.log(LogStatus.INFO, "Click on Order No " +vr.stckno);
							}
						
						//check status of order no
						String status = w.findElement(By.xpath("//td[text()='"+vr.stckno+"']/following-sibling::td[6]")).getText();
						System.out.println("status is "+status);	
						
						///click change order
						w.findElement(By.partialLinkText("CHANGE ORDER")).click();
								
						String check = "Error";
						if(w.getPageSource().contains(check))
						{
							String popup1 = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
							log.log(LogStatus.ERROR, "" +popup1);
							System.out.println("Order no " +vr.stckno+ " cannot be modified as status is " +status);
							w.findElement(By.partialLinkText("BACK")).click();
						}
						else
						{														
							//update stock no
							j=j+1;
							if(w.findElement(By.id("stk_qty")).isEnabled())
							{
							w.findElement(By.id("stk_qty")).clear();
							vr.quantity = (int) scu.getRow(i).getCell(j).getNumericCellValue();  	
							System.out.println("data1 " +vr.quantity);
							w.findElement(By.id("stk_qty")).sendKeys(String.valueOf(vr.quantity));
							log.log(LogStatus.INFO, "Modified Quantity is : " +vr.quantity);
							//wb.close();
							
							///modify order
							vr.RM = ((vr.quantity * vr.roundoffpr ) / vr.multiple);
							System.out.println("RM " +vr.RM);
							vr.roundoff = (Math.round(vr.RM*100.0)/100.0);
							System.out.println("roundoff " +vr.roundoff);
							}
							else
							{
								w.findElement(By.id("stk_lot")).clear();
								vr.quantity = (int) scu.getRow(i).getCell(j).getNumericCellValue();  	
								w.findElement(By.id("stk_lot")).sendKeys(String.valueOf(vr.quantity));
								log.log(LogStatus.INFO, "Modified No of lots : " +vr.quantity);
								w.findElement(By.id("stk_price")).click();
								String qt1 = w.findElement(By.id("stk_qty")).getAttribute("value");
								vr.quantity = Integer.parseInt(qt1);
								log.log(LogStatus.INFO, "Quantity is : " +vr.quantity);
								
								///modify order
								vr.RM = ((vr.quantity * vr.roundoffpr ) / vr.multiple);
								vr.roundoff = (Math.round(vr.RM*100.0)/100.0);
								System.out.println("roundoff " +vr.roundoff);
							}
							
							vr.rw = scu.getRow(i);
							vr.cell = vr.rw.createCell(variables.cwmactutil);
							vr.cell.setCellValue(vr.roundoff);
							
							///click on chnge order
							w.findElement(By.partialLinkText("CHANGE ORDER")).click();
							log.log(LogStatus.INFO, "Click on Change Order");	
							
							//cnfrm
							w.findElement(By.partialLinkText("Confirm")).click();
							log.log(LogStatus.INFO, "Click Confirm");
							System.out.println("Order no " +vr.stckno+ " is modified successfully");
							log.log(LogStatus.INFO, "Order no " +vr.stckno+ " is modified successfully");
														
							log.log(LogStatus.INFO, "Modified required margin is " +vr.roundoff);
							
							Actions act15 = new Actions(w);
							act15.moveToElement(w.findElement(By.linkText("Funds"))).click().perform();	
							WebElement Sublink15 = w.findElement(By.linkText("Equity/Derivatives"));
							Sublink15.click();
										
							//capture utilisation total
							w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).click();
							
							//if(scrip1.equals(scripchk))
							if(w.getPageSource().contains(vr.scrip1))
							{
								///get util amount of specific scrip				
								vr.total1 = w.findElement(By.xpath("//td[text()='"+vr.scrip1+"']/following-sibling::td[6]")).getText();
								System.out.println("total1 " +vr.total1);
								vr.tot = Double.parseDouble(vr.total1);
								System.out.println("cap amt " +vr.tot);
								
								//log.log(LogStatus.INFO, "Margin Utilised Cash & FNO " +tot);
																								
								vr.rw = scu.getRow(i);
								vr.cell = vr.rw.createCell(variables.cwmutil);
								vr.cell.setCellValue(vr.tot);
								
								//Mcmprmaxutil();
									double result;
									
									vr.action = scu.getRow(i).getCell(8).getStringCellValue();
										
									if(vr.action.equals("BUY"))
									{
										bma = vr.roundoff;
										System.out.println("buy action total" +bma);
										if(vr.tot>=bma)
										{
											bma=vr.tot;
											System.out.println("buy action total" +vr.tot);
										}
									}
									else
									{
										sma = vr.roundoff;
										System.out.println("sell action total" +sma);
										if(vr.tot>=sma)
										{
											sma=vr.tot;
											System.out.println("sell action total" +vr.tot);
										}
									}
									
									if(bma>=sma)
									{
										result = bma;
										System.out.println("buyamt "+bma);
										bma=0;
									}
									else
									{
										result = sma;
										System.out.println("sellamt "+sma);
										sma=0;
									}
									
									vr.rw = scu.getRow(i);
									vr.cell = vr.rw.createCell(variables.cwmmaxutil);
									vr.cell.setCellValue(result);
									
							}
												
							}
												
						}
						
						else
						{
							System.out.println("Order no not generated as unable to place order.");
							log.log(LogStatus.ERROR, "Order no not generated as unable to modify order.");
						}
						
					  }
						}
					  	break;
	  			}	  									
				  	
			}
				else
				{
				}
				
	  	}
		wb.write(Of);
		Of.close();
	}
	
	  
	@Test(priority=2)
	public void CancelOrder() throws Exception 
	{
	  	FileOutputStream Of = new FileOutputStream(srcfile);
	  			
		  	for(i=5;i<rc+1;i++)
		  	{
	  		Row row = scu.getRow(i);
	  		String flag = scu.getRow(i).getCell(0).getStringCellValue();
				
				if(flag.equals("Y"))
				{
					vr.forreportscripname=scu.getRow(i).getCell(1).getStringCellValue();
				  	System.out.println("reportscripname "+vr.forreportscripname);
				  	log = report1.startTest("Cancel " +vr.forreportscripname+ " " +actiontype+ " Order");
	  		
	  				for(j=15;j<row.getLastCellNum();)
	  				{	
	  					vr.scrip1 = scu.getRow(i).getCell(5).getStringCellValue();
	  					  ///check for any error
						  String amocheck1 = "Error";
						  if(w.getPageSource().contains(amocheck1))
						  	{
							  	String capturemsg = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
						  		log.log(LogStatus.ERROR, "" +capturemsg);
						  		
						   	}
						  else
						  {
							///click watchlist>>place order
							Actions act2 = new Actions(w);
							act2.moveToElement(w.findElement(By.linkText("Reports"))).click().perform();	
							WebElement Sublink2 = w.findElement(By.linkText("Order Status"));
							Sublink2.click();
							log.log(LogStatus.INFO, "Click on Reports >> Order Status");
							
							vr.stckno = scu.getRow(i).getCell(j).getStringCellValue();
							System.out.println("cancel stk no " +vr.stckno);
																					
							//check stkno
							if(vr.stckno.equals("null")|| vr.stckno.equals(""))
							{
								System.out.println("Order no not generated as unable to place order.");
								log.log(LogStatus.ERROR, "Order no not generated as unable to cancel order.");
							}
							else
							{
							if(vr.stckno !=null &&  !vr.stckno.isEmpty())
							{
							if(w.getPageSource().contains(vr.stckno))
							{
							w.findElement(By.xpath(".//input[starts-with(@value,'"+vr.stckno+"')]")).click();
							log.log(LogStatus.INFO, "Click on Order No " +vr.stckno);
							}
							
							///click on cancel
							w.findElement(By.partialLinkText("CANCEL ORDER")).click();
							log.log(LogStatus.INFO, "Click on Cancel");
																					
							///check for error
							String trade = "Can not cancel Order.";
							//String msg = "--";
							if(w.getPageSource().contains(trade)) //|| w.getPageSource().equals("--"))
							{
								String popup2 = w.findElement(By.xpath("//*[@id=\'container\']/div/form/table[1]/tbody/tr[2]/td[7]")).getText();
								System.out.println("status is "+popup2);
								log.log(LogStatus.ERROR, "" +popup2);				
								w.findElement(By.partialLinkText("Back")).click();	
								System.out.println("Order no "+vr.stckno+ " cannot be cancelled as status is " +vr.status);
								
								vr.newamt = vr.cap;
								System.out.println("newamt " +vr.newamt);
								
								vr.rw = scu.getRow(i);
								vr.cell = vr.rw.createCell(variables.cwcutil);
								vr.cell.setCellValue(vr.cap);
							}
							else
							{
								//cnfrm
								w.findElement(By.partialLinkText("Confirm")).click();
								log.log(LogStatus.INFO, "Click on Confirm");
								System.out.println("Order no "+vr.stckno+ " is cancelled successfully");
								log.log(LogStatus.INFO, "Order no " +vr.stckno+ " is cancelled successfully");
										
								vr.newamt = vr.cap;
								System.out.println("newamt " +vr.newamt);
																
								///call capture total
								j=5;
								System.out.println("j new "+j);
								capturetotal();
								
								vr.rw = scu.getRow(i);
								vr.cell = vr.rw.createCell(variables.cwcutil);
								vr.cell.setCellValue(vr.cap);
								
								Double canamt = 0.0;
								System.out.println("can amt " +Math.abs(canamt));
								log.log(LogStatus.INFO, "Cancelled Required Margin is " +Math.abs(canamt));
								
								vr.rw = scu.getRow(i);
								vr.cell = vr.rw.createCell(variables.cwcactutil);
								vr.cell.setCellValue(Math.abs(canamt));
								
								//Ccmprmaxutil();
									double result;
									
									vr.action = scu.getRow(i).getCell(8).getStringCellValue();
										
									if(vr.action.equals("BUY"))
									{
										bma = canamt;
										System.out.println("buy action total" +bma);
										if(vr.cap>=bma)
										{
											bma=vr.cap;
											System.out.println("buy action total" +vr.cap);
										}
									}
									else
									{
										sma = canamt;
										System.out.println("sell action total" +sma);
										if(vr.cap>=sma)
										{
											sma=vr.cap;
											System.out.println("sell action total" +vr.cap);
										}
									}
									
									if(bma>=sma)
									{
										result = bma;
										System.out.println("buyamt "+bma);
										bma=0;
									}
									else
									{
										result = sma;
										System.out.println("sellamt "+sma);
										sma=0;
									}		
									vr.rw = scu.getRow(i);
									vr.cell = vr.rw.createCell(variables.cwcmaxutil);
									vr.cell.setCellValue(result);
								
								vr.rw = scu.getRow(i);
								vr.cell = vr.rw.createCell(variables.cwclastutil);
								vr.cell.setCellValue(vr.cap);
									
							}
						  
						  }
							else
							{
								System.out.println("Order no not generated as unable to place order.");
								log.log(LogStatus.ERROR, "Order no not generated as unable to cancel order.");
							}
							}
						  
						 }
						
						break;
	  						  			
					}
	  				
				}
				
				else
	  			{
	  			}	  		
	}
		wb.write(Of);
		Of.close();
	}	
	
	@AfterTest
	public void aftertest()
	{
		//w.close();
	}
	
	@Test(priority=4002)
	public void nouse()
	  {
		  report1.endTest(log);
		  report1.flush();
		  w.get("D:\\testfile\\cmprmaxutil.html");
		  w.get(srcfile);
	  }
}
