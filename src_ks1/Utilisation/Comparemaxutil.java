package Utilisation;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.testng.annotations.Test;

import com.library.variables;
import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.LogStatus;

import Main.BSOrderXL;

public class Comparemaxutil extends BSOrderXL
{
	public static String capture;
	public static String scripname;
	public static String popup1;
	public static Double cap;
	public static Double matchround;
	public static Double roundoff;
	public static Double roundoff1;
	public static Double newroundoff;
	public static Double price;
	public static Double newamt;
	public static Double roundoffpr;
	public static Double tot;
	public static Double RM;
	public static Double total;
	public static String stc;
	public static float multiple;
	public static int quantity;
	public static String sd;
	public static String scrip1;
	public static String action;
	public static String pr1;
	//public String stckno;
	public String Status1;
	public String status;
	public static int j;
	public static int i;
	public static int k;
	public static int rc;
	public static Row rw;
	public static Cell cell;
	public static String combinemng;
	public static String scripvalmng;
	public static HSSFSheet scu;
	public static String watchlistname;
	public static String chk;
	public static WebElement ele;
	public static String scrip;
	public static String combine;
	public static String getexactexpdate;
	public static double buymaxamt = 0;
	public static double sellmaxamt=0;
	public static String forreportscripname;
	public static List<WebElement> stckno;
	public static String arr[];
	public static double bma;
	public static double sma;

	
	//String arr[] = new String[stckno.size()];
	ArrayList<String> mylist = new ArrayList<String>();
	
	//com.utilisation.ManageWatchlist cm = new com.utilisation.ManageWatchlist();
	com.library.functions fn = new com.library.functions();
	
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
		watchlistname = scu.getRow(1).getCell(5).getStringCellValue();  	
		w.findElement(By.xpath("//*[@id=\'container\']/div/form/table[1]/tbody/tr/td[2]/input")).sendKeys(watchlistname);
		log.log(LogStatus.INFO, "Enter Watchlist Name : " +watchlistname);
		
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
			if(w.getPageSource().contains(watchlistname))
			{
				List<WebElement> radio = w.findElements(By.xpath("//input[@name='id' and @type = 'radio']"));
				for(int i=0;i<radio.size();i++)
				{
					ele = radio.get(i);
					String value=ele.getAttribute("id");
					chk =w.findElement(By.xpath("//div[@id='container']//a[contains(text(), '"+watchlistname+"')]")).getText();
				}		
				if(chk.contains(watchlistname))
				{
					ele.click();
				}
				
				w.findElement(By.partialLinkText("DELETE")).click();
				log.log(LogStatus.INFO, "Existed Watchlist deleted successfully");
				w.findElement(By.partialLinkText("ADD")).click();
				log.log(LogStatus.INFO, "Add watchlist");
				w.findElement(By.xpath("//*[@id=\'container\']/div/form/table[1]/tbody/tr/td[2]/input")).sendKeys(watchlistname);
				log.log(LogStatus.INFO, "Enter Watchlist Name : " +watchlistname);
				w.findElement(By.partialLinkText("Submit")).click();
				log.log(LogStatus.INFO, "Click on Submit");
				w.findElement(By.xpath("//div[@id='container']//a[contains(text(), '"+watchlistname+"')]")).click();
				log.log(LogStatus.INFO, "Select " +watchlistname+ " to add scrips");
				
			}
		}
		else
			{
				///select added watchlist
				w.findElement(By.xpath("//div[@id='container']//a[contains(text(), '"+watchlistname+"')]")).click();
				log.log(LogStatus.INFO, "Watchlist " +watchlistname+ " gets added successfully " );
				
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
		  			scrip = scu.getRow(i).getCell(j).getStringCellValue();
		  			w.findElement(By.id("search_list")).sendKeys(scrip);
		  			
		  			j=j+1;
		  			String ME = scu.getRow(i).getCell(j).getStringCellValue();
		  			w.findElement(By.id("me")).sendKeys(ME);
		  				  				  			
		  			j=j+1;
		  			String IT = scu.getRow(i).getCell(j).getStringCellValue();
		  			w.findElement(By.id("inst_type")).sendKeys(IT);
		  				  			
		  			w.findElement(By.partialLinkText("SUBMIT")).click();	  								  						  								
	
		  			//to check if scrip exist
		  			if(w.getPageSource().contains(scrip.toUpperCase()))
		        	{	  				
		  				fn.getdate();
		  				
		        	}
		  			else
		  			{
		  				w.findElement(By.partialLinkText("CANCEL")).click();
			  			log.log(LogStatus.INFO, "No matching stocks found for the Scrip " +scrip.toUpperCase());
			  			
			  			j=j+1;
			  			rw = scu.getRow(i);
			  			cell = rw.createCell(j);
			  			cell.setCellValue("null");
			  			
			  			j=j+1;		  			  			
			  			rw = scu.getRow(i);
			  			cell = rw.createCell(j);
			  			cell.setCellValue("null");	
			  			
		  				break;
		  			}
		  			
		  			if(ME.equals("NSE") || ME.equals("BSE"))
		  			{
			  			String space = " ";
			  			combine = scrip.toUpperCase().concat(space).concat(IT).concat(space).concat(ME);
			  			System.out.println("combine "+combine);
		  			}
		  			else
		  			{
		  				String space = " ";
			  			combine = scrip.toUpperCase().concat(space).concat(IT).concat(space).concat(ME).concat(space).concat(getexactexpdate);
			  			System.out.println("combine "+combine);
		  			}
		  			
		  			j=j+1;
		  			if(ME.equals("NSE-Derivatives"))
		  			{
		  				ME="NSE";
		  				String space = " ";
			  			combine = scrip.toUpperCase().concat(space).concat(IT).concat(space).concat(ME).concat(space).concat(getexactexpdate);

		  			}
		  			if(ME.equals("BSE-Derivatives"))
		  			{
		  				ME="BSE";
		  				String space = " ";
			  			combine = scrip.toUpperCase().concat(space).concat(IT).concat(space).concat(ME).concat(space).concat(getexactexpdate);

		  			}
		  					  			
		  			rw = scu.getRow(i);
		  			cell = rw.createCell(j);
		  			cell.setCellValue(combine);
		  					  			
		  			j=j+1;		  			  			
		  			rw = scu.getRow(i);
		  			cell = rw.createCell(j);
		  			cell.setCellValue(fn.b);		  			  			
		  			
		  			log.log(LogStatus.INFO, "Scrip "+scrip.toUpperCase()+ " got added.");
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
				
		scrip1 = scu.getRow(i).getCell(j).getStringCellValue();
		System.out.println("scrip1 "+scrip1);
		if(!w.getPageSource().contains(checkmsg))
		{
			if(w.getPageSource().contains(scrip1))
			{
			System.out.println("in loop scrip1 " +scrip1);
			String capture = w.findElement(By.xpath("//td[text()='"+scrip1+"']/following-sibling::td[6]")).getText();
			cap = Double.parseDouble(capture);
			System.out.println("first/cap " +cap);
			}
			else
			{
				cap = 0.0;
				System.out.println("cap of capture " +cap);
			}
			rw = scu.getRow(i);
			cell = rw.createCell(variables.cwputil);
			cell.setCellValue(cap);
		}
		else
		{
			System.out.println("Out of loop");
		}
				
	}
	
	public void Pcmprmaxutil() throws Exception
	{				
		if(action.equals("BUY"))
		{
			//capbuyamt = roundoff;
			bma = buymaxamt + roundoff;
			//System.out.println("buy action total" +buymaxamt);
			System.out.println("buy action total" +bma);
		}
		else
		{
			//capsellamt = roundoff;
			sma = sellmaxamt + roundoff;
			//System.out.println("sell action total" +sellmaxamt);
			System.out.println("buy action total" +sma);
		}
		
		if(bma>=sma)
		{
			System.out.println("buyamt "+bma);
			rw = scu.getRow(i);
			cell = rw.createCell(variables.cwpmaxutil);
			cell.setCellValue(bma);
		}
		else
		{
			System.out.println("sellamt "+sma);
			rw = scu.getRow(i);
			cell = rw.createCell(variables.cwpmaxutil);
			cell.setCellValue(sma);
		}
	}
	
	public void Mcmprmaxutil() throws Exception
	{
		if(action.equals("BUY"))
		{
			//capbuyamt = roundoff;
			bma = buymaxamt + roundoff1;
			System.out.println("buy action total" +bma);
			//System.out.println("buy action total" +bma);
		}
		else
		{
			//capsellamt = roundoff;
			sma = sellmaxamt + roundoff1;
			System.out.println("sell action total" +sma);
			//System.out.println("buy action total" +sma);
		}
		
		if(bma>=sma)
		{
			System.out.println("buyamt "+bma);
			rw = scu.getRow(i);
			cell = rw.createCell(variables.cwmmaxutil);
			cell.setCellValue(bma);
		}
		else
		{
			System.out.println("sellamt "+sma);
			rw = scu.getRow(i);
			cell = rw.createCell(variables.cwmmaxutil);
			cell.setCellValue(sma);
		}
	}
	
	@Test(priority=0)
	public void PlaceOrder() throws Exception 
	  {
		FileInputStream f = new FileInputStream(srcfile);
		wb = new HSSFWorkbook(f);
	  	FileOutputStream Of = new FileOutputStream(srcfile);

		scu = wb.getSheet("Compareutil");
				
		///report creation
		report1 = new ExtentReports("D:\\Saleema\\cmprutil.html");
		log = report1.startTest("Manage Watchlist");
	  	
		////call login function
	  	Login();
	  	
	  	///call manage watchlist
	  	manage();
	  	
	  	///call add scrip
	  	addloopscrips();
	  		  		  	
	  	///get multiple value from excel for utilization caln
	  	
	  	//int rc = scu.getLastRowNum() - scu.getFirstRowNum();
	  	
	  	//System.out.println("rc "+rc);
	  	//String watchlistname = scu.getRow(1).getCell(5).getStringCellValue();
	  	
	  	for(i=5;i<rc+1;i++)
	  	{
	  		Row row = scu.getRow(i);
	  		 
	  		String flag = scu.getRow(i).getCell(0).getStringCellValue();
				
				if(flag.equals("Y"))
				{
					forreportscripname=scu.getRow(i).getCell(1).getStringCellValue();
				  	System.out.println("reportscripname "+forreportscripname);
				  	log = report1.startTest("Place " +forreportscripname+ " Order");
	  		
	  				for(j=5;j<row.getLastCellNum();)
	  				{	
					  	///call capture total function to capture utilisation value
	  					capturetotal();
						
					  	///click watchlist>>place order
						Actions act = new Actions(w);
						act.moveToElement(w.findElement(By.linkText("Watchlist"))).click().perform();	
						WebElement Sublink = w.findElement(By.linkText("My Watchlists"));
						Sublink.click();
						log.log(LogStatus.INFO, "Click on Watchlist >> My Watchlists");
						
						///watchlistname									
						System.out.println("watchlistname "+watchlistname);
						w.findElement(By.xpath("//div[@id='container']//a[contains(text(), '"+watchlistname+"')]")).click();
						
						///select scrip
						j=j+1;
						String xcelscrip = scu.getRow(i).getCell(j).getStringCellValue();
						System.out.println("data1 " +xcelscrip);
						w.findElement(By.xpath(".//input[starts-with(@value,'"+xcelscrip+"')]")).click();
						
						j=j+1;
						String getexccode = w.findElement(By.xpath(".//td[input[starts-with(@value,'"+xcelscrip+"')]]/following-sibling::td[3]")).getText();
						System.out.println("getexchange code "+getexccode);
						
						rw = scu.getRow(i);
						cell = rw.createCell(j);
						cell.setCellValue(getexccode);
												
						////click on place order
						w.findElement(By.partialLinkText("PLACE ORDER")).click();
						log.log(LogStatus.INFO, "Click on Place Order");
						
						///get scripname
						scripname = w.findElement(By.id("stk_name")).getAttribute("value");
						log.log(LogStatus.INFO, "Selected Scrip is " +scripname);
						
						///select buy/sell
						j=j+1;
						action = scu.getRow(i).getCell(j).getStringCellValue();
						System.out.println("boolean " +action);
						w.findElement(By.xpath(".//input[@value='"+action+"']")).click();
						log.log(LogStatus.INFO, "Selected Value is " +action);
											
						////enter quantity
						j=j+1;
						int quantity = (int) scu.getRow(i).getCell(j).getNumericCellValue();  	
						System.out.println("data1 " +quantity);
						//wb.close();
						
						if(w.findElement(By.id("stk_qty")).isEnabled())
						{
							w.findElement(By.id("stk_qty")).sendKeys(String.valueOf(quantity));
							log.log(LogStatus.INFO, "Entered Quantity is: " +quantity);
							
							///multiple
							j=j+1;
							multiple = (float) scu.getRow(i).getCell(j).getNumericCellValue();  	
							System.out.println("multiple " +multiple);
							
							///required margin calculation
							j=j+1;
							int perc = (int) scu.getRow(i).getCell(j).getNumericCellValue();
							pr1 = w.findElement(By.id("stk_price")).getAttribute("value");
							w.findElement(By.id("stk_price")).clear();
							price = (Double.parseDouble(pr1)*perc)/100;
							double chngestrngpr = Double.parseDouble(pr1);
							if(action.equals("BUY"))
							{
								total = chngestrngpr-price;
							}
							else
							{
								total = chngestrngpr+price;
							}
							roundoffpr = (double) (Math.round(total));
							w.findElement(By.id("stk_price")).sendKeys(String.valueOf(roundoffpr));
							log.log(LogStatus.INFO, "Actual price is " +pr1);
							log.log(LogStatus.INFO, "Modified price is " +roundoffpr);
							RM = ((quantity * roundoffpr ) / multiple);
							roundoff = (Math.round(RM*100.0)/100.0);
							log.log(LogStatus.INFO, "Required Margin is " +roundoff);
							System.out.println("roundoff " +roundoff);
						}
						else
						{
							w.findElement(By.id("stk_lot")).sendKeys(String.valueOf(quantity));
							log.log(LogStatus.INFO, "Entered No of lots is: " +quantity);
							
							///multiple
							j=j+1;
							multiple = (float) scu.getRow(i).getCell(j).getNumericCellValue();  	
							System.out.println("multiple  " +multiple);
							
							///required margin calculation
							j=j+1;
							int perc = (int) scu.getRow(i).getCell(j).getNumericCellValue();
							pr1 = w.findElement(By.id("stk_price")).getAttribute("value");
							w.findElement(By.id("stk_price")).clear();
							price = (Double.parseDouble(pr1)*perc)/100;
							double chngestrngpr = Double.parseDouble(pr1);
							if(action.equals("BUY"))
							{
								total = chngestrngpr-price;
							}
							else
							{
								total = chngestrngpr+price;
							}
							roundoffpr = (double) (Math.round(total));
							w.findElement(By.id("stk_price")).sendKeys(String.valueOf(roundoffpr));
							log.log(LogStatus.INFO, "Actual price is " +pr1);
							log.log(LogStatus.INFO, "Modified price is " +roundoffpr);
							String qty = w.findElement(By.id("stk_qty")).getAttribute("value");
							log.log(LogStatus.INFO, "Quantity is: " +qty);
							quantity = Integer.parseInt(qty);
							RM = ((quantity * roundoffpr ) / multiple);
							roundoff = (Math.round(RM*100.0)/100.0);
							log.log(LogStatus.INFO, "Required Margin is " +roundoff);
							System.out.println("roundoff " +roundoff);
						}
						
						///write actual price value to excel
						j=j+1;
						rw = scu.getRow(i);
						cell = rw.createCell(j);
						cell.setCellValue(pr1);
						
						///write modified price value to excel
						j=j+1;
						rw = scu.getRow(i);
						cell = rw.createCell(j);
						cell.setCellValue(roundoffpr);
																														
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
							popup1 = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
							System.err.println("Error displayed as " +popup1);
							log.log(LogStatus.ERROR, "" +popup1);
							if(w.getPageSource().contains("BACK"))
							{
								w.findElement(By.partialLinkText("BACK")).click();
								stckno = null;
								String arr[]=new String[stckno.size()];
								arr[i]=null;
							}
							else
							{
								System.out.println("...");
							}
						}
						else
						{	
							Thread.sleep(1000);
							////capture order no
							stckno= w.findElements(By.xpath("//*[@id='container']/div/p/span"));
							String arr[]=new String[stckno.size()];
					        System.out.println("stckno "+stckno.size());
					 	    for(int k=0;k<rc+1;)
					 	    {
					 	    	arr[k]=stckno.get(k).getText();
					 	    	//stckno = w.findElement(By.xpath("//*[@id='container']/div/p/span")).getText();
					 	    	System.out.println("stckno...... " +arr[k]);
					 	    	log.log(LogStatus.INFO, "Order No is " +arr[k]);
					 	    	
					 	    	mylist.add(arr[k]);
					 	    	System.out.println("mylist size "+mylist.size());
					 	    	break;
					 	    }													
							///click funds >>equity/derivatives
							Actions act1 = new Actions(w);
							act1.moveToElement(w.findElement(By.linkText("Funds"))).click().perform();	
							WebElement Sublink1 = w.findElement(By.linkText("Equity/Derivatives"));
							Sublink1.click();
							//log.log(LogStatus.INFO, "Click on Funds >> Equity/Derivatives");
							
							//click margin utilised cash&FNO
							w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).click();
							
							if(w.getPageSource().contains(scrip1))
							{
								///get util amount of specific scrip				
								String total = w.findElement(By.xpath("//td[text()='"+scrip1+"']/following-sibling::td[6]")).getText();
								System.out.println("total " +total);
								tot = Double.parseDouble(total);
								System.out.println("cap amt " +tot);
								Double match = tot - cap;
								System.out.println("match " +match);
								matchround = Math.round(match*100.0)/100.0;
								System.out.println("f amt" +matchround);
								//log.log(LogStatus.INFO, "Margin Utilised Cash & FNO " +matchround);
																								
								rw = scu.getRow(i);
								cell = rw.createCell(variables.cwpactutil);
								cell.setCellValue(roundoff);
								
								Pcmprmaxutil();

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
	  	int k;
	  	
	  	//for(k=0;k<mylist.size();k++)
	  	//{
	  		for(i=5;i<rc+1;i++)
		  	{	
	  		Row row = scu.getRow(i);
	  		 
	  		String flag = scu.getRow(i).getCell(0).getStringCellValue();
				
				if(flag.equals("Y"))
				{
					forreportscripname=scu.getRow(i).getCell(1).getStringCellValue();
				  	System.out.println("reportscripname "+forreportscripname);
				  	log = report1.startTest("Change " +forreportscripname+ " Order");
	  		
				  	for(k=0;k<mylist.size();k++)
  					{
				  	
	  				for(j=15;j<row.getLastCellNum();)
	  				{
	  					//for(k=0;k<mylist.size();k++)
	  					//{
						//capturetotal();
						
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
											  	
						//check stkno
						String abc = mylist.get(k);
						//arr[k]=mylist.get(k);
						System.out.println("manage stk no " +abc);
						if(abc !=null &&  !abc.isEmpty())
						{
							if(w.getPageSource().contains(abc))
							{
								w.findElement(By.xpath(".//input[starts-with(@value,'"+abc+"')]")).click();
								log.log(LogStatus.INFO, "Click on Order No " +abc);
							}
						
						Thread.sleep(1000);
						
						//check status of order no
						String status = w.findElement(By.xpath("//td[text()='"+abc+"']/following-sibling::td[6]")).getText();
						System.out.println("status is "+status);	
						
						///click change order
						w.findElement(By.partialLinkText("CHANGE ORDER")).click();
								
						String check = "Error";
						if(w.getPageSource().contains(check))
						{
							String popup1 = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
							log.log(LogStatus.ERROR, "" +popup1);
							//w.findElement(By.partialLinkText("BACK")).click();
							System.out.println("Order no " +stckno+ " cannot be modified as status is " +status);

						}
						else
						{														
							Double newcap = matchround;
							System.out.println("newcap " +newcap);							
																				
							//update stock no
							if(w.findElement(By.id("stk_qty")).isEnabled())
							{
							w.findElement(By.id("stk_qty")).clear();
							quantity = (int) scu.getRow(i).getCell(j).getNumericCellValue();  	
							System.out.println("data1 " +quantity);
							w.findElement(By.id("stk_qty")).sendKeys(String.valueOf(quantity));
							log.log(LogStatus.INFO, "Modified Quantity is : " +quantity);
							//wb.close();
							
							///modify order
							RM = ((quantity * roundoffpr ) / multiple);
							System.out.println("RM " +RM);
							roundoff1 = (Math.round(RM*100.0)/100.0);
							System.out.println("roundoff " +roundoff1);
							}
							else
							{
								w.findElement(By.id("stk_lot")).clear();
								quantity = (int) scu.getRow(i).getCell(j).getNumericCellValue();  	
								w.findElement(By.id("stk_lot")).sendKeys(String.valueOf(quantity));
								log.log(LogStatus.INFO, "Modified No of lots : " +quantity);
																
								w.findElement(By.id("stk_price")).click();
								
								String qt1 = w.findElement(By.id("stk_qty")).getAttribute("value");
								quantity = Integer.parseInt(qt1);
								log.log(LogStatus.INFO, "Quantity is : " +quantity);
								///modify order
								RM = ((quantity * roundoffpr ) / multiple);
								roundoff1 = (Math.round(RM*100.0)/100.0);
								System.out.println("roundoff " +roundoff1);
							}
							
							rw = scu.getRow(i);
							cell = rw.createCell(variables.cwmactutil);
							cell.setCellValue(roundoff1);
							
							///click on chnge order
							w.findElement(By.partialLinkText("CHANGE ORDER")).click();
							log.log(LogStatus.INFO, "Click on Change Order");	
							
							//cnfrm
							w.findElement(By.partialLinkText("Confirm")).click();
							log.log(LogStatus.INFO, "Click Confirm");
							System.out.println("Order no " +abc+ " is modified successfully");
							log.log(LogStatus.INFO, "Order no " +abc+ " is modified successfully");
														
							log.log(LogStatus.INFO, "Modified required margin is " +roundoff1);
							
							Actions act15 = new Actions(w);
							act15.moveToElement(w.findElement(By.linkText("Funds"))).click().perform();	
							WebElement Sublink15 = w.findElement(By.linkText("Equity/Derivatives"));
							Sublink15.click();
										
							//capture utilisation total
							w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).click();
							
							if(w.getPageSource().contains(scrip1))
							{
								///get util amount of specific scrip				
								String total = w.findElement(By.xpath("//td[text()='"+scrip1+"']/following-sibling::td[6]")).getText();
								System.out.println("total " +total);
								tot = Double.parseDouble(total);
								System.out.println("cap amt " +tot);
								
								//log.log(LogStatus.INFO, "Margin Utilised Cash & FNO " +tot);
																								
								rw = scu.getRow(i);
								cell = rw.createCell(variables.cwmutil);
								cell.setCellValue(tot);
								Mcmprmaxutil();

							}
												
							}
												
						}
						else
						{
							System.out.println("Order no not generated as unable to place order.");
							log.log(LogStatus.ERROR, "Order no not generated as " +popup1);
						}
				  	
					  }
	  				//}
					  	break;
	  			}
	  				//break;
				}	
				  	//break;
			}
				else
				{
				}
				//break;
	  	}
		  	
	  		//}
			  	
		wb.write(Of);
		Of.close();
	}
	
/*	  
	@Test(priority=2)
	public void CancelOrder() throws Exception 
	{
	  	FileOutputStream Of = new FileOutputStream(srcfile);
	  		  	
	  	int k;
	  	
		for(k=0;k<stckno.size();k++)
	  	{
		  	for(i=5;i<rc+1;i++)
		  	{
	  		Row row = scu.getRow(i);
	  		 
	  		String flag = scu.getRow(i).getCell(0).getStringCellValue();
				
				if(flag.equals("Y"))
				{
					forreportscripname=scu.getRow(i).getCell(1).getStringCellValue();
				  	System.out.println("reportscripname "+forreportscripname);
				  	log = report1.startTest("Cancel " +forreportscripname+ " Order");
	  		
	  				for(j=5;j<row.getLastCellNum();)
	  				{	
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
							
							//check stkno
							if(abc !=null &&  !abc.isEmpty())
							{
							if(w.getPageSource().contains(abc))
							{
							w.findElement(By.xpath(".//input[starts-with(@value,'"+abc+"')]")).click();
							log.log(LogStatus.INFO, "Click on Order No " +abc);
							}
							
							///click on cancel
							w.findElement(By.partialLinkText("CANCEL ORDER")).click();
							log.log(LogStatus.INFO, "Click on Cancel");
																					
							///check for error
							String trade = "Can not cancel Order.";
							if(w.getPageSource().contains(trade))
							{
								String popup2 = w.findElement(By.xpath("//*[@id=\'container\']/div/form/table[1]/tbody/tr[2]/td[7]")).getText();
								System.out.println("status is "+popup2);
								log.log(LogStatus.ERROR, "" +popup2);				
								w.findElement(By.partialLinkText("Back")).click();	
								System.out.println("Order no "+abc+ " cannot be cancelled as status is " +status);
								
								newamt = cap;
								System.out.println("newamt " +newamt);
								
								rw = scu.getRow(i);
								cell = rw.createCell(variables.cwcutil);
								cell.setCellValue(newamt);
							}
							else
							{
								//cnfrm
								w.findElement(By.partialLinkText("Confirm")).click();
								log.log(LogStatus.INFO, "Click on Confirm");
								System.out.println("Order no "+abc+ " is cancelled successfully");
								log.log(LogStatus.INFO, "Order no " +abc+ " is cancelled successfully");
										
								newamt = cap;
								rw = scu.getRow(i);
								cell = rw.createCell(variables.cwcutil);
								cell.setCellValue(newamt);
								
								///call capture total
								j=5;
								System.out.println("j new "+j);
								capturetotal();
								Double canamt = newamt - cap;
								System.out.println("can amt " +canamt);
								log.log(LogStatus.INFO, "Cancelled Required Margin is " +canamt);
								cmprmaxutil();
									
							}
						  
						  }
							else
							{
								System.out.println("Order no not generated as unable to place order.");
								log.log(LogStatus.ERROR, "Order no not generated as " +popup1);
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
	}
		wb.write(Of);
		Of.close();
	}				
		
	*/
	
	@Test(priority=45)
	public void nouse()
	  {
		  report1.endTest(log);
		  report1.flush();
		  w.get("D:\\Saleema\\cmprutil.html");
	  }
}
