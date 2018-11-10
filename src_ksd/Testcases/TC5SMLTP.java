package Testcases;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.testng.Assert;
import org.testng.TestListenerAdapter;
import org.testng.TestNG;
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

public class TC5SMLTP extends BSOrderXL
{
	  int i,j,rc,c;
	  public static WebElement ele;
	  public float perc,SLOprice;
	  public double chngestrngpr;
	  public static String status,highprice,lowprice,formatted,scrip,getscrip,qty;
	  int column,column1,column2;
	  
	  com.library.functions fn = new com.library.functions();
	  com.library.variables vr = new com.library.variables();
	  
	  public static void main(String[] args) 
	  {
		  TestListenerAdapter tla = new TestListenerAdapter();
		    TestNG testng = new TestNG();
		     Class[] classes = new Class[]
		     {
		    	 TC5SMLTP.class     
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
			vr.watchlistname = vr.s.getRow(1).getCell(5).getStringCellValue();  	
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
			int i,j;
			
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
			  				fn.getdate();
			  				
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
			  					  			
			  			vr.rw = vr.s.getRow(i);
			  			vr.cell = vr.rw.createCell(j);
			  			vr.cell.setCellValue(vr.combine);
			  					  			
			  			j=j+1;		  			  			
			  			vr.rw = vr.s.getRow(i);
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
			String checkmsg = "No Record Found";
					
			vr.scrip1 = vr.s.getRow(i).getCell(j).getStringCellValue();
			System.out.println("scrip1 "+vr.scrip1);
			
			if(vr.scrip1 !=null &&  !vr.scrip1.isEmpty())
			{
				if(!w.getPageSource().contains(checkmsg))
				{
					if(w.getPageSource().contains(vr.scrip1))
					{
						System.out.println("in loop scrip1 " +vr.scrip1);
						//List<WebElement> myList= w.findElements(By.xpath("//tr[@class ='tinside']"));
						List<WebElement> myList= w.findElements(By.xpath("//div[@id ='container']/div/div[1]/table/tbody/tr[@class ='tinside']"));
						System.out.println("list1 size "+myList.size());
				        String arr[]=new String[myList.size()];
				        String arr1[]=new String[myList.size()];
				        c=2;
				        
				 	    for(int i=0;i<myList.size();i++)
				 	    {					 	    	//tr[@class ='tinside']
				 	    	arr[i]= myList.get(i).findElement(By.xpath("//*[@id='container']/div/div/table/tbody/tr[" + (c)+ "]/td[1]")).getText();				 	    	
							System.out.println("arr of i "+arr[i]);
																			
							arr1[i]= myList.get(i).findElement(By.xpath("//*[@id='container']/div/div/table/tbody/tr[" + (c)+ "]/td[3]")).getText();
							System.out.println("arr of i "+arr1[i]);
							
							if(arr[i].startsWith(vr.scrip1))
							{
								if(arr1[i].equals("-"))
								{
									String capture = w.findElement(By.xpath("//*[@id='container']/div/div/table/tbody/tr[" + (c)+ "]/td[7]")).getText();
									vr.cap = Double.parseDouble(capture);
									System.out.println("first/cap " +vr.cap);
									//break;
								}
								else
								{
									vr.cap = 0.0;
								}																	
							}
							else
							{
								vr.cap = 0.0;
							}
							c=c+1;
							System.out.println("..");							
				 	    }				 	    
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
		
	  
	  @Test(priority=0)
	  public void PlaceSMorder() throws Exception
	  {
		  	FileInputStream f = new FileInputStream(vr.srcfile);
			vr.wb = new HSSFWorkbook(f);
		  	
			vr.s = vr.wb.getSheet("SMLtp");
						
			///report creation
			report1 = new ExtentReports("D:\\testfile\\supermultiple.html");
			log = report1.startTest("Manage Watchlist");
			
			////call login function
		  	Login();
						
			///call manage watchlist
		  	manage();
		  		  	
		  	///call add scrip
		  	addloopscrips();
		  	
			rc = vr.s.getLastRowNum() - vr.s.getFirstRowNum();
		  	System.out.println("rc "+rc);
		  	
		  	///place SM order
			
		  	for(i=5;i<rc+1;i++)
		  	{
		  		Row row = vr.s.getRow(i);
		  		
		  		String flag = vr.s.getRow(i).getCell(0).getStringCellValue();
					
					if(flag.equals("Y"))
					{
				  		vr.forreportscripname= vr.s.getRow(i).getCell(1).getStringCellValue();
					  	System.out.println("reportscripname "+vr.forreportscripname);
					  	log = report1.startTest("Place " +vr.forreportscripname+ " Order");
					  	
		  				for(j=5;j<row.getLastCellNum();)
		  				{					
////////////				///call capturetotal
							capturetotal();
							
							vr.rw = vr.s.getRow(i);
							vr.cell = vr.rw.createCell(variables.smutil);
							vr.cell.setCellValue(vr.cap);
							
		  			  		vr.scrip1 = vr.s.getRow(i).getCell(5).getStringCellValue();						  			  					
		  			  		
						  	///click watchlist>>place order
							Actions act = new Actions(w);
							act.moveToElement(w.findElement(By.linkText("Watchlist"))).click().perform();	
							WebElement Sublink = w.findElement(By.linkText("My Watchlists"));
							Sublink.click();
							log.log(LogStatus.INFO, "Click on Watchlist >> My Watchlists");
							
							///watchlistname									
							System.out.println("watchlistname "+vr.watchlistname);
							w.findElement(By.xpath("//div[@id='container']//a[contains(text(), '"+vr.watchlistname+"')]")).click();
									
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
							
							///click on view details
							w.findElement(By.partialLinkText("VIEW DETAILS")).click();
							log.log(LogStatus.INFO, "Click on View details");
							
							if(vr.forreportscripname.equals("Derivative") || vr.forreportscripname.equals("Currency"))
							{
								///capture high price range
								highprice = w.findElement(By.xpath("//*[@id='container']/div/table[2]/tbody/tr/td[1]/table/tbody/tr[13]/td[2]")).getText();
								System.out.println("high price "+highprice);
								double hp = Double.parseDouble(highprice);
								
								///capture low price range
								lowprice = w.findElement(By.xpath("//*[@id='container']/div/table[2]/tbody/tr/td[1]/table/tbody/tr[13]/td[4]")).getText();
								System.out.println("low price "+lowprice);
								double lp = Double.parseDouble(lowprice);
							}
							else
							{
								///capture high price range
								highprice = w.findElement(By.xpath("//*[@id='container']/div/table[2]/tbody/tr/td[1]/table/tbody/tr[10]/td[2]")).getText();
								System.out.println("high price "+highprice);
								double hp = Double.parseDouble(highprice);
								
								///capture low price range
								lowprice = w.findElement(By.xpath("//*[@id='container']/div/table[2]/tbody/tr/td[1]/table/tbody/tr[10]/td[4]")).getText();
								System.out.println("low price "+lowprice);
								double lp = Double.parseDouble(lowprice);
							}
							
							log.log(LogStatus.INFO, "High Price Range "+highprice+ " and Low Price Range "+lowprice);
							w.navigate().back();
							
							///select scrip
							w.findElement(By.xpath(".//input[starts-with(@value,'"+xcelscrip+"')]")).click();
							
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
							{	//////////for cash
								w.findElement(By.id("stk_qty")).sendKeys(String.valueOf(vr.quantity));
								log.log(LogStatus.INFO, "Entered Quantity is: " +vr.quantity);		
								
								///multiple
								j=j+1;
								vr.multiple = (float) vr.s.getRow(i).getCell(j).getNumericCellValue();  	
								System.out.println("multiple " +vr.multiple);
								
								///actual price calculation for cash
								j=j+1;
								perc = (float) vr.s.getRow(i).getCell(j).getNumericCellValue();
								System.out.println("perc " +perc);
								vr.pr1 = w.findElement(By.id("stk_price")).getAttribute("value");
								System.out.println("pr1........."+vr.pr1);
								w.findElement(By.id("stk_price")).clear();
								vr.price = (Double.parseDouble(vr.pr1)*perc)/100;
								chngestrngpr = Double.parseDouble(vr.pr1);
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
								if(perc==0.0)
								{
									chngestrngpr = 0.0;
								}
////////////////////////////////////////////////////////////////////////////////////////////////////////
								
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
							{	///////////for FNO
								w.findElement(By.id("stk_lot")).sendKeys(String.valueOf(vr.quantity));
								log.log(LogStatus.INFO, "Entered No of lots is: " +vr.quantity);
								
								///multiple
								j=j+1;
								vr.multiple = (float) vr.s.getRow(i).getCell(j).getNumericCellValue();  	
								System.out.println("multiple " +vr.multiple);
								
								///actual price calculation for FNO
								j=j+1;
								perc = (float) vr.s.getRow(i).getCell(j).getNumericCellValue();
								vr.pr1 = w.findElement(By.id("stk_price")).getAttribute("value");
								w.findElement(By.id("stk_price")).clear();
								vr.price = (Double.parseDouble(vr.pr1)*perc)/100;
								chngestrngpr = Double.parseDouble(vr.pr1);
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
								if(perc==0.0)
								{
									chngestrngpr = 0.0;
								}
////////////////////////////////////////////////////////////////////////////////////////////////////////
								
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
								System.out.println("quantity "+vr.quantity);
							}
							
							///write actual price value to excel
							j=j+1;
							vr.rw = vr.s.getRow(i);
							vr.cell = vr.rw.createCell(j);
							vr.cell.setCellValue(vr.pr1);
							
							///write modified price value to excel
							j=j+1;
							vr.rw = vr.s.getRow(i);
							vr.cell = vr.rw.createCell(j);
							vr.cell.setCellValue(vr.roundoffpr);
							
							///select order type
							Select OT = new Select(w.findElement(By.id("del")));
							OT.selectByVisibleText("Super Multiple Plus");
							log.log(LogStatus.INFO, "Select Order type : Super Multiple Plus");
							
							if(vr.forreportscripname.equals("Currency"))
							{
								j=j+1;
								vr.RM = ((vr.quantity * vr.roundoffpr ) / vr.multiple);
								System.out.println("RM "+vr.RM);
								vr.reqmargin = (Math.round(vr.RM*100.0)/100.0);
								System.out.println("currency SM" +vr.reqmargin);
							}
							else
							{	///////for cash and derivative
								///enter SLO price
								j=j+1;
								SLOprice = (float) vr.s.getRow(i).getCell(j).getNumericCellValue();
								vr.HLprice = (vr.roundoffpr*SLOprice)/100;
								System.out.println("hlprice "+vr.HLprice);
								if(vr.action.equals("BUY"))
								{
									vr.calHLprice = vr.roundoffpr-vr.HLprice;
								}
								else
								{
									vr.calHLprice = vr.roundoffpr+vr.HLprice;
								}
								//vr.calHLprice = vr.roundoffpr-vr.HLprice;
								vr.roundoff1 = (double) (Math.round(vr.calHLprice));
								w.findElement(By.id("stk_t_price")).sendKeys(String.valueOf(vr.roundoff1));
								System.out.println("slo price " +vr.roundoff1);
								log.log(LogStatus.INFO, "SLO Trigger Price is " +vr.roundoff1);
								
								DecimalFormat format = new DecimalFormat("0.00");
						        formatted = format.format(vr.roundoff1);
						        System.out.println(formatted);
						        
						        ///cal req margin
								if(vr.action.equals("BUY"))
								{
							        vr.caldiff = vr.roundoffpr-vr.roundoff1;						        
								}
								else
								{
							        vr.caldiff = vr.roundoff1-vr.roundoffpr;						        
								}
						        //vr.caldiff = vr.roundoffpr-vr.roundoff1;						        
								float catgry = (float) vr.s.getRow(i).getCell(variables.category).getNumericCellValue();
								float SLOcat = (float) vr.s.getRow(i).getCell(variables.diffcat).getNumericCellValue();
								vr.calLtp = (vr.roundoffpr *catgry)/100;
								//vr.calLtp = (Math.round(vr.calLtp1*100.0)/100.0);
								vr.adddiffLtp = vr.caldiff+vr.calLtp;
								//vr.roundoff4 = (Math.round(vr.adddiffLtp*100.0)/100.0);
								System.out.println("quantiyti "+vr.quantity);
								Double rm = vr.quantity * vr.adddiffLtp;
								vr.reqmargin = (Math.round(rm*100.0)/100.0);
								System.out.println("diff calc " +vr.caldiff+ " catgry " +catgry+ " callltp "+vr.calLtp);
								System.out.println("add diff calc " +vr.adddiffLtp);
								//log.log(LogStatus.INFO, "Required Margin is " +vr.reqmargin);
								System.out.println("req margin "+vr.reqmargin);
								if(vr.action.equals("BUY"))
								{	/////to check if SLO trigger price is < then trade price
									System.out.println("buy req margin "+vr.reqmargin);
								}
								else
								{	/////to check if SLO trigger price is > then trade price							
									vr.addcat = catgry + SLOcat ;
									System.out.println("add cat "+vr.addcat);
									vr.calLtp1 = (vr.roundoffpr * vr.addcat)/100;
									vr.rm1 = vr.quantity * vr.calLtp1;
									System.out.println("rm1 "+vr.rm1);	
									if(vr.reqmargin>vr.rm1)
									{
										vr.reqmargin = vr.reqmargin;
										System.out.println("sell req mar1 "+vr.reqmargin);
									}
									else
									{
										vr.reqmargin = vr.rm1;
										System.out.println("sell req mar2 "+vr.reqmargin);
									}
								}
							}
							
							///place order
							w.findElement(By.partialLinkText("PLACE ORDER")).click();
							log.log(LogStatus.INFO, "Click on Place Order button");
							
							///alert
							w.switchTo().alert().accept();
							log.log(LogStatus.INFO, "Click OK on alert");
							
							///for cash
							String capture = "Trigger Price["+formatted+"] is out of Price range ["+lowprice+"] and ["+highprice+"]";
							System.out.println("capture msg "+capture);
							if(w.getPageSource().contains(capture))
							{
								String pricecap = w.findElement(By.xpath("//*[@id='container']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
								System.out.println("msg "+pricecap);
								String[] a = pricecap.split("\\[");
								String[] b = a[2].split("]");
								String[] c = a[3].split("]");
								System.out.println("b " +b[0]);
								System.out.println("c " +c[0]);
								
								if(highprice.equals(c[0]) && lowprice.equals(b[0]))
								{
									log.log(LogStatus.INFO, "Trigger high price is " +lowprice+ " and Trigger low price is " +highprice);
									log.log(LogStatus.INFO, "Trigger price range is equal");
									
									///write slo high price
									j=j+1;
									vr.rw = vr.s.getRow(i);
									vr.cell = vr.rw.createCell(j);
									vr.cell.setCellValue(highprice);
									
									///write slo low price
									j=j+1;
									vr.rw = vr.s.getRow(i);
									vr.cell = vr.rw.createCell(j);
									vr.cell.setCellValue(lowprice);									
								}
								else
								{
									log.log(LogStatus.INFO, "Trigger high price is " +lowprice+ " and Trigger low price is " +highprice);
									log.log(LogStatus.INFO, "Trigger price range is not equal");
									
									///write slo high price
									j=j+1;
									vr.rw = vr.s.getRow(i);
									vr.cell = vr.rw.createCell(j);
									vr.cell.setCellValue(highprice);
									
									///write slo low price
									j=j+1;
									vr.rw = vr.s.getRow(i);
									vr.cell = vr.rw.createCell(j);
									vr.cell.setCellValue(lowprice);
								}
								w.findElement(By.partialLinkText("BACK")).click();
								break;
							}		
							
							///for derivative							 
							String capdermsg = "Trigger Price is not within the price range";// 1235.00 and 1365.00";
							System.out.println("capture msg "+capdermsg);	
							
							///check for any error						    
							String popup = "Error";
							if (w.getPageSource().contains(popup))
							{
								String expln = "Explanation";
								if(w.getPageSource().contains(expln))
								{
									vr.popup1 = w.findElement(By.xpath("//*[@id='container']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
									String[] a1 = vr.popup1.split("of", 0);
									System.out.println("a111 "+a1[0]);
									if(a1[0].contains(capdermsg))
									{										
										String[] a = vr.popup1.split(" ");
										System.out.println("lp " +a[9]);
										System.out.println("hp " +a[11]);
										
										if(highprice.equals(a[11]) && lowprice.equals(a[9]))
										{
											log.log(LogStatus.INFO, "Trigger high price is " +a[9]+ " and Trigger low price is " +a[11]);
											log.log(LogStatus.INFO, "Trigger price range is equal");
											
											///write slo high price
											j=j+1;
											vr.rw = vr.s.getRow(i);
											vr.cell = vr.rw.createCell(j);
											vr.cell.setCellValue(highprice);
											
											///write slo low price
											j=j+1;
											vr.rw = vr.s.getRow(i);
											vr.cell = vr.rw.createCell(j);
											vr.cell.setCellValue(lowprice);									
										}
										else
										{
											log.log(LogStatus.INFO, "Trigger high price is " +a[9]+ " and Trigger low price is " +a[11]);
											log.log(LogStatus.INFO, "Trigger price range is not equal");
											
											///write slo high price
											j=j+1;
											vr.rw = vr.s.getRow(i);
											vr.cell = vr.rw.createCell(j);
											vr.cell.setCellValue(highprice);
											
											///write slo low price
											j=j+1;
											vr.rw = vr.s.getRow(i);
											vr.cell = vr.rw.createCell(j);
											vr.cell.setCellValue(lowprice);
										}
									}
								}
								else
								{
									vr.popup1 = w.findElement(By.xpath("//*[@id='container']/div/table/tbody/tr[2]/td/table/tbody/tr[1]/td")).getText();

								}
								System.err.println("Error displayed as " +vr.popup1);
								log.log(LogStatus.ERROR, "" +vr.popup1);
								if(w.getPageSource().contains("BACK"))
								{
									w.findElement(By.partialLinkText("BACK")).click();											
									break;
								}
								else
								{
									System.out.println("...");
								}
								vr.rw = vr.s.getRow(i);
								vr.cell = vr.rw.createCell(variables.smstatus);
								vr.cell.setCellValue(vr.popup1);
							}
							else
							{	
								log.log(LogStatus.INFO, "Required Margin is " +vr.reqmargin);
								
								vr.rw = vr.s.getRow(i);
								vr.cell = vr.rw.createCell(variables.smreqmar);
								vr.cell.setCellValue(vr.reqmargin);
								
								Thread.sleep(1000);
								////capture order no
								vr.stckno = w.findElement(By.xpath("//*[@id='container']/div/p/span")).getText();
						 	    System.out.println("stckno...... " +vr.stckno);
						 	    log.log(LogStatus.INFO, "Order No is " +vr.stckno);
						 	    
								j=j+3;
								vr.rw = vr.s.getRow(i);
								vr.cell = vr.rw.createCell(j);
								vr.cell.setCellValue(vr.stckno);	
								
//////////////////////////////////////////////////////////////////////////////////////////////////
								////////////////place order calculation//////////////////////
								
								///click funds >>equity/derivatives
								Actions act1 = new Actions(w);
								act1.moveToElement(w.findElement(By.linkText("Funds"))).click().perform();	
								WebElement Sublink1 = w.findElement(By.linkText("Equity/Derivatives"));
								Sublink1.click();
								log.log(LogStatus.INFO, "Click on Funds >> Equity/Derivatives");
								
								//click margin utilised cash&FNO
								w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).click();
								
								System.out.println("selected scrip is " +vr.scrip1);
								if(w.getPageSource().contains(vr.scrip1))
								{									
									///get util amount of specific scrip
									System.out.println("in loop scrip1 " +vr.scrip1);
									List<WebElement> myList= w.findElements(By.xpath("//tr[@class ='tinside']"));
									System.out.println("list size "+myList.size());
					
							        String arr[]=new String[myList.size()];
							        String arr1[]=new String[myList.size()];
							        c=2;
							 	    for(int x=0;x<myList.size();x++)
							 	    {					 	    	//tr[@class ='tinside']
							 	    	arr[x]= myList.get(x).findElement(By.xpath("//*[@id='container']/div/div/table/tbody/tr[" + (c)+ "]/td[1]")).getText();				 	    	
										System.out.println("arr of x "+arr[x]);																						
										//arr1[i]= myList.get(i).findElement(By.xpath("//td[text()='"+arr[i]+"']/following-sibling::td[2]")).getText();
										arr1[x]= myList.get(x).findElement(By.xpath("//*[@id='container']/div/div/table/tbody/tr[" + (c)+ "]/td[3]")).getText();
										System.out.println("arr of x "+arr1[x]);
										
										if(arr[x].startsWith(vr.scrip1))
										{
											if(arr1[x].equals("-"))
											{
												String total = w.findElement(By.xpath("//*[@id='container']/div/div/table/tbody/tr[" + (c)+ "]/td[7]")).getText();
												System.out.println("total " +total);
												vr.tot = Double.parseDouble(total);
												System.out.println("cap amt " +vr.tot);
												
												Double match = vr.tot - vr.cap;
												System.out.println("match " +match);
												vr.matchround = Math.round(match*100.0)/100.0;
												System.out.println("f amt" +vr.matchround);
												log.log(LogStatus.INFO, "Margin Utilised " +vr.matchround);
												
												vr.rw = vr.s.getRow(i);
												vr.cell = vr.rw.createCell(variables.smmutil);
												vr.cell.setCellValue(vr.tot);
												
												vr.rw = vr.s.getRow(i);
												vr.cell = vr.rw.createCell(variables.smmarutil);
												vr.cell.setCellValue(vr.matchround);
												
												if(vr.matchround.equals(vr.reqmargin))
												{
													///click on cancel
													w.findElement(By.partialLinkText("Cancel")).click();
													log.log(LogStatus.INFO, "Utilized -- As Required Margin and Margin Utilised are Equal");
													
													vr.rw = vr.s.getRow(i);
													vr.cell = vr.rw.createCell(variables.smstatus);
													vr.cell.setCellValue("Required and utilised margin are equal");
												}
												else
												{
													log.log(LogStatus.ERROR, "Not Utilized -- As Required Margin and Margin Utilised are Unequal");
													
													vr.rw = vr.s.getRow(i);
													vr.cell = vr.rw.createCell(variables.smstatus);
													vr.cell.setCellValue("Required and utilised margin are unequal");
												}
											}								
										}
										c=c+1;
										System.out.println("..");
							 	    }		
								}
							}
//////////////////////////////////////////////////////////////////////////////////////////////////							
							///manage order
							log = report1.startTest("Change " +vr.forreportscripname+ " Order");													
							
							///to check for any error
							String amocheck = "Error";
						  	if(w.getPageSource().contains(amocheck))
						  	{
						  		String capturemsg = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
						  		log.log(LogStatus.ERROR, "" +capturemsg);
						  		
						  		vr.rw = vr.s.getRow(i);
						  		vr.cell = vr.rw.createCell(variables.smstatus);
						  		vr.cell.setCellValue(capturemsg);
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
							System.out.println("manage stk no " +vr.stckno);
							if(vr.stckno !=null &&  !vr.stckno.isEmpty())
							{
							if(w.getPageSource().contains(vr.stckno))
							{
							w.findElement(By.xpath(".//input[starts-with(@value,'"+vr.stckno+"')]")).click();
							log.log(LogStatus.INFO, "Click on Order No " +vr.stckno);
							}
							Thread.sleep(1000);
							//check status of order no
							String status = w.findElement(By.xpath("//td[text()='"+vr.stckno+"']/following-sibling::td[6]")).getText();
							if(status.equals("AMO") || status.equals("OPN"))
							{
											
							///click change order
							w.findElement(By.partialLinkText("CHANGE ORDER")).click();
									
							String check = "Error";
							if(w.getPageSource().contains(check))
							{
								String popup1 = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
								log.log(LogStatus.ERROR, "" +popup1);
								if(w.getPageSource().contains("BACK"))
								{
									w.findElement(By.partialLinkText("BACK")).click();
									break;
								}
								else
								{
									System.out.println("...");
								}
								System.out.println("Order no " +vr.stckno+ " cannot be modified as status is " +status);
								
								vr.rw = vr.s.getRow(i);
								vr.cell = vr.rw.createCell(variables.smstatus);
								vr.cell.setCellValue(popup1);

							}
							else
							{	/////change order price to 0 and chk low trigger price error then trade
								Double newcap = vr.matchround;
								System.out.println("newcap " +newcap);
								
								//update price
								w.findElement(By.id("stk_price")).clear();
								w.findElement(By.id("stk_price")).sendKeys("0");
								log.log(LogStatus.INFO, "Modify Price as 0");
																	
								w.findElement(By.id("stk_t_price")).clear();
								
								///to chk low trigger price message
								vr.HLprice = (vr.roundoffpr*SLOprice)/100;
								System.out.println("hlprice "+vr.HLprice);
								if(vr.action=="BUY")
								{
									vr.calHLprice = vr.roundoffpr+vr.HLprice;
								}
								else
								{
									vr.calHLprice = vr.roundoffpr-vr.HLprice;
								}
								//vr.calHLprice = vr.roundoffpr+vr.HLprice;
								vr.roundoff2 = (double) (Math.round(vr.calHLprice));
								System.out.println("roundoff2 "+vr.roundoff2);
								w.findElement(By.id("stk_t_price")).sendKeys(String.valueOf(vr.roundoff2));
								log.log(LogStatus.INFO, "Enter SLO Trigger price : "+vr.roundoff1);
								
								w.findElement(By.partialLinkText("CHANGE ORDER")).click();
								String sloerr = "Error";
								if(w.getPageSource().contains(sloerr))
								{
									///capture low price
									String capturelp = w.findElement(By.xpath("//*[@id='container']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
									String getlp = capturelp.split(" ")[9];
									log.log(LogStatus.ERROR, "" +capturelp);								
									
									w.findElement(By.partialLinkText("BACK")).click();
									
									w.findElement(By.id("stk_t_price")).clear();
									w.findElement(By.id("stk_t_price")).sendKeys("0");
									log.log(LogStatus.INFO, "SLO Trigger Price is : 0");
									
									///click on chnge order
									w.findElement(By.partialLinkText("CHANGE ORDER")).click();
									log.log(LogStatus.INFO, "Click on Change Order");
								}
								else
								{
									System.out.println("No trigger message");
								}									
								
								//cnfrm
								w.findElement(By.partialLinkText("Confirm")).click();
								log.log(LogStatus.INFO, "Click Confirm");
								
								System.out.println("Order no " +vr.stckno+ " is modified successfully");
								log.log(LogStatus.INFO, "Order no " +vr.stckno+ " is modified successfully");																
									
								w.findElement(By.partialLinkText("Order Status")).click();
								
								w.findElement(By.xpath(".//input[starts-with(@value,'"+vr.stckno+"')]")).click();
								String getLTP = w.findElement(By.xpath("//td[text()='"+vr.stckno+"']/following-sibling::td[3]")).getText();
								System.out.println("getLTP "+getLTP);
								double capLTP = Double.parseDouble(getLTP);
								
								////cal modify req margin
								if(vr.forreportscripname.equals("Currency"))
								{
									vr.RM = ((vr.quantity * capLTP ) / vr.multiple);
									System.out.println("RM " +vr.RM);
									vr.modreqmargin = (Math.round(vr.RM*100.0)/100.0);
									System.out.println("Mod req margin is " +vr.modreqmargin);
								}
								else
								{	////////modify order calculation////////////////////////
									if(vr.action.equals("BUY"))
									{
										vr.caldiff =  capLTP - vr.roundoff1;
									}
									else
									{
										vr.caldiff =  vr.roundoff1 - capLTP;
									}
							        //vr.caldiff =  capLTP - vr.roundoff1;						        
									float catgry = (float) vr.s.getRow(i).getCell(variables.category).getNumericCellValue();
									float SLOcat = (float) vr.s.getRow(i).getCell(variables.diffcat).getNumericCellValue();
									vr.calLtp = (capLTP *catgry)/100;
									//vr.calLtp = (Math.round(vr.calLtp1*100.0)/100.0);
									vr.adddiffLtp = vr.caldiff+vr.calLtp;
									//vr.roundoff3 = (Math.round(vr.adddiffLtp*100.0)/100.0);
									System.out.println("quantiyti "+vr.quantity);
									Double mrm = vr.quantity * vr.adddiffLtp;
									vr.modreqmargin = (Math.round(mrm*100.0)/100.0);
									System.out.println("diff calc " +vr.caldiff+ " catgry " +catgry+ " callltp "+vr.calLtp);
									System.out.println("add diff calc " +vr.adddiffLtp);
									//log.log(LogStatus.INFO, "Required Margin is " +vr.modreqmargin);
									System.out.println("Mod req margin is "+vr.modreqmargin);
									
									if(vr.action.equals("BUY"))
									{
										System.out.println("mod buy req margin "+vr.modreqmargin);
									}
									else
									{	///////check modify util if SLO price > then trade price										
										vr.addcat = catgry + SLOcat ;
										System.out.println("add cat "+vr.addcat);
										vr.calLtp2 = (vr.roundoffpr * vr.addcat)/100;
										System.out.println("callltp2 "+vr.calLtp2);
										vr.mrm1 = vr.quantity * vr.calLtp2;
										System.out.println("mrm1 "+vr.mrm1);	
										if(vr.modreqmargin>vr.mrm1)
										{
											vr.modreqmargin = vr.modreqmargin;
											System.out.println("sell req mar1 "+vr.modreqmargin);
										}
										else
										{
											vr.modreqmargin = vr.mrm1;
											System.out.println("sell req mar1 "+vr.modreqmargin);
										}
									}
									
								}								
								vr.rw = vr.s.getRow(i);
								vr.cell = vr.rw.createCell(variables.smmreqmar);
								vr.cell.setCellValue(vr.modreqmargin);
								
								Actions act15 = new Actions(w);
								act15.moveToElement(w.findElement(By.linkText("Funds"))).click().perform();	
								WebElement Sublink15 = w.findElement(By.linkText("Equity/Derivatives"));
								Sublink15.click();
											
								//capture utilisation total
								w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).click();
								
								if(w.getPageSource().contains(vr.scrip1))
								{
									///get util amount of specific scrip	
									System.out.println("in loop scrip1 " +vr.scrip1);
									List<WebElement> myList= w.findElements(By.xpath("//tr[@class ='tinside']"));
							        String arr[]=new String[myList.size()];
							        String arr1[]=new String[myList.size()];
							        c=2;
							 	    for(int y=0;y<myList.size();y++)
							 	    {					 	    	//tr[@class ='tinside']
							 	    	arr[y]= myList.get(y).findElement(By.xpath("//*[@id='container']/div/div/table/tbody/tr[" + (c)+ "]/td[1]")).getText();				 	    	
										System.out.println("arr of y "+arr[y]);
																						
										arr1[y]= myList.get(y).findElement(By.xpath("//*[@id='container']/div/div/table/tbody/tr[" + (c)+ "]/td[3]")).getText();
										System.out.println("arr of y "+arr1[y]);
										
										if(arr[y].startsWith(vr.scrip1))
										{
											if(arr1[y].equals("-"))
											{
												String total = w.findElement(By.xpath("//*[@id='container']/div/div/table/tbody/tr[" + (c)+ "]/td[7]")).getText();
												System.out.println("mod order total " +total);
												vr.tot = Double.parseDouble(total);
												System.out.println("mod order cap amt " +vr.tot);
												Double match = vr.tot - vr.cap;
												System.out.println("mod order match " +match);
												vr.matchround = Math.round(match*100.0)/100.0;
												System.out.println("mod order mar util matchround" +vr.matchround);
												
												vr.rw = vr.s.getRow(i);
												vr.cell = vr.rw.createCell(variables.smcutil);
												vr.cell.setCellValue(vr.tot);
												
												if(vr.matchround.equals(vr.modreqmargin))
												{
													///click on cancel
													w.findElement(By.partialLinkText("Cancel")).click();
													System.out.println("equal ");
													log.log(LogStatus.INFO, "Utilized -- As Required Margin and Margin Utilised are Equal");
																										
												}
												else
												{
													log.log(LogStatus.ERROR, "Not Utilized -- As Required Margin and Margin Utilised are Unequal");
																								
												}
											}								
										}
										c=c+1;
										System.out.println("..");
							 	    }
								}
								
								}
							}
							else
							{
								w.findElement(By.partialLinkText("CHANGE ORDER")).click();
								log.log(LogStatus.INFO, "Click on Change Order");	
								String chkerror = "Error";
								if(w.getPageSource().contains(chkerror))
								{
									String capmoderror = w.findElement(By.xpath("//*[@id=\'container\']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
									log.log(LogStatus.ERROR, "" +capmoderror);
									
									vr.rw = vr.s.getRow(i);
									vr.cell = vr.rw.createCell(variables.smstatus);
									vr.cell.setCellValue(capmoderror);
									
									w.findElement(By.partialLinkText("BACK")).click();
									//w.findElement(By.xpath(".//input[starts-with(@value,'"+vr.stckno+"')]")).click();
								}
							  }
							}
							else
							{
								System.out.println("Order no not generated as unable to place order.");
								log.log(LogStatus.ERROR, "Order no not generated as " +vr.popup1);
								vr.rw = vr.s.getRow(i);
								vr.cell = vr.rw.createCell(variables.smstatus);
								vr.cell.setCellValue(vr.popup1);								
							}
					  	
						  	}	
						  	
							log = report1.startTest("Cancel " +vr.forreportscripname+ " SLO Order");
						
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////							
							///cancel SLO and check message
							
							Actions act1 = new Actions(w);
							act1.moveToElement(w.findElement(By.linkText("Reports"))).click().perform();	
							WebElement Sublink1 = w.findElement(By.linkText("Order Status"));
							Sublink1.click();						  							
							
							Long stk = Long.valueOf(vr.stckno);
							Long sno = stk + 2;
							String cnvtsno = sno.toString(); ///convert sno from long to string
							///write SLO order no to excel
							j=j+1;
							vr.rw = vr.s.getRow(i);
							vr.cell = vr.rw.createCell(j);
							vr.cell.setCellValue(cnvtsno);
							
							w.findElement(By.xpath(".//input[starts-with(@value,'"+sno+"')]")).click();
							log.log(LogStatus.INFO, "Click on SLO Order No " +sno);
							w.findElement(By.partialLinkText("CANCEL ORDER")).click();
							log.log(LogStatus.INFO, "Click on Cancel Order");
							
							///check for SLO error
							String trade = "Can not cancel Order.";
							if(w.getPageSource().contains(trade))
							{
								String popup2 = w.findElement(By.xpath("//*[@id='container']/div/form/table[1]/tbody/tr[2]/td[7]")).getText();
								System.out.println("status is "+popup2);
								log.log(LogStatus.ERROR, "" +popup2);	
								
								vr.rw = vr.s.getRow(i);
								vr.cell = vr.rw.createCell(variables.smstatus);
								vr.cell.setCellValue(popup2);
								
								w.findElement(By.partialLinkText("Back")).click();	
								System.out.println("Order no "+vr.stckno+ " cannot be cancelled as status is " +vr.status);	
								break;
							}
							else
							{
								w.findElement(By.partialLinkText("Confirm")).click();
								String cancelmsg = "Can not Cancel Super Multiple Leg 2 Order";
								String cmsg = "Can not Cancel Leg 2 Order";
								if(w.getPageSource().contains(cancelmsg) || w.getPageSource().contains(cmsg))
								{
									log.log(LogStatus.ERROR, "" +cancelmsg);
								}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////								
								///modify SLO qty and check message
								Actions act2 = new Actions(w);
								act2.moveToElement(w.findElement(By.linkText("Reports"))).click().perform();	
								WebElement Sublink2 = w.findElement(By.linkText("Order Status"));
								Sublink2.click();
								log.log(LogStatus.INFO, "Click on Reports >> Order Status");
								
								w.findElement(By.xpath(".//input[starts-with(@value,'"+sno+"')]")).click();
								log.log(LogStatus.INFO, "Click on SLO Order No " +sno);
								
								w.findElement(By.partialLinkText("CHANGE ORDER")).click();
								System.out.println("j.........."+j);
								
								if(w.findElement(By.id("stk_qty")).isEnabled())
								{
									j=j+1;
									w.findElement(By.id("stk_qty")).clear();
									vr.modqty = (int) vr.s.getRow(i).getCell(j).getNumericCellValue();
									System.out.println("mod qty " +vr.modqty);
									w.findElement(By.id("stk_qty")).sendKeys(String.valueOf(vr.modqty));
									log.log(LogStatus.INFO, "Modified Quantity is : " +vr.modqty);									
								}
								else
								{
									j=j+1;
									w.findElement(By.id("stk_lot")).clear();
									vr.modqty = (int) vr.s.getRow(i).getCell(j).getNumericCellValue();  	
									System.out.println("mod qty " +vr.modqty);
									w.findElement(By.id("stk_lot")).sendKeys(String.valueOf(vr.modqty));
									log.log(LogStatus.INFO, "Modified Quantity is : " +vr.modqty);																		
								}								
																										
								w.findElement(By.partialLinkText("CHANGE ORDER")).click();
								log.log(LogStatus.INFO, "Click on Change Order");
								
								String moderr = "Error";
								if(w.getPageSource().contains(moderr))
								{
									String capmoderr= w.findElement(By.xpath("//*[@id='container']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
									log.log(LogStatus.ERROR, "" +capmoderr);
									w.findElement(By.partialLinkText("BACK")).click();
									break;
								}
								else
								{																	
									w.findElement(By.partialLinkText("Confirm")).click();
									
									String qtymsg = w.findElement(By.xpath("//*[@id='container']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
									log.log(LogStatus.ERROR, "" +qtymsg);	
								}
							}
							
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////							
							///place square off order
							log = report1.startTest("Place " +vr.forreportscripname+ " Square Off Order");
							
							Actions act3 = new Actions(w);
							act3.moveToElement(w.findElement(By.linkText("Reports"))).click().perform();	
							WebElement Sublink3 = w.findElement(By.linkText("Open Position"));
							Sublink3.click();
							log.log(LogStatus.INFO, "Click on Reports >> Open Position");	
							
							System.out.println("xcelscrip "+xcelscrip);
							String[] x = xcelscrip.split("\\(" ,0);
							System.out.println("x[0] "+x[0]);
							List<WebElement> myList= w.findElements(By.xpath(".//td[input[@type = 'checkbox']]"));
							List<WebElement> myListele = w.findElements(By.xpath("//input[@name='chk' and @type = 'checkbox']"));
							for(int i=0;i<myList.size();i++)
						 	{										
					 	    	WebElement ele;
					 	    	ele=myListele.get(i);
								String actualscrip = ele.getAttribute("value");
								System.out.println("actualscrip "+actualscrip);								
								//System.out.println("qty "+vr.quantity);
								//String fvalue = x[0].concat(","+vr.quantity).concat(",99");
								//System.out.println("print fvalue "+fvalue);
								
								if(actualscrip.startsWith(x[0]))
								{									
									WebElement mytable = w.findElement(By.xpath("//div[@id='container']/div/form/div/table/tbody"));
							    	//To locate rows of table. 
							    	List < WebElement > rows_table = mytable.findElements(By.tagName("tr"));
							    	//To calculate no of rows In table.
							    	int rows_count = rows_table.size();
							    	//Loop will execute till the last row of table.
							    	for (int row1 = 2; row1 < rows_count;row1++) 
							    	{
							    	    //To locate columns(cells) of that specific row.
							    	    List < WebElement > Columns_row = rows_table.get(row1).findElements(By.tagName("td"));
							    	    //To calculate no of columns (cells). In that specific row.
							    	    int columns_count = Columns_row.size();			
							    	    System.out.println("Number of cells In Row " + row1 + " are " + columns_count);
							    	    
							    	    //Loop will execute till the last cell of that specific row.
							    	    for (column = 1; column < columns_count;)//column++) 
							    	    {
							    	        // To retrieve text from that specific cell.
							    	        scrip = Columns_row.get(column).getText();
							    	        System.out.println("Cell Value of row number " + row1 + " and column number " + column + " Is " + scrip);
							    	        							    	        
							    	        if(scrip.contains(vr.scrip1))
							    	        {
							    	        	column1=column+2;
							    	        	column2=column+3;
							        	        scrip = Columns_row.get(column1).getText();
							        	        qty = Columns_row.get(column2).getText();
							        	        String fvalue1 = x[0].concat(","+qty).concat(",99");
							        	        System.out.println("fvalue1 "+fvalue1);
							        	        if(actualscrip.startsWith(fvalue1))
							        	        {
								    	            if(scrip.contains("Super Multiple Plus"))
									    	        {
								    	            	System.out.println("Super Multiple Plus");
								    	            	ele.click();
								    	            	log.log(LogStatus.INFO, "Select scrip : " +vr.scripname);
								    	            	break;
									    	        }	
								    	            							    	            
								    	            else// if(scrip.contains("Normal"))
									    	        {
									    	        	System.out.println("other type");
									    	        	//ele.click();
									    	        	break;
									    	        }
							        	        }							    	            
							    	        }
							    	        else
							    	        {
							    	        	break;
							    	        }	
							    	        break;
							    	    }							    	    
							    	}	
							    	//break;
								}								
						 	}
							
							//if(ele.isSelected())
							{	
								w.findElement(By.partialLinkText("PLACE ORDER")).click();
								///price value
								w.findElement(By.id("price00")).clear();
								vr.price = (Double.parseDouble(vr.pr1)*perc)/100;
								chngestrngpr = Double.parseDouble(vr.pr1);
								vr.total = chngestrngpr+vr.price;
								vr.roundoffpr = (double) (Math.round(vr.total));
								
								////////enter order price more than ltp
								w.findElement(By.id("price00")).sendKeys(String.valueOf(vr.roundoffpr));
								log.log(LogStatus.INFO, "Enter price "+vr.pr1);
								
								w.findElement(By.partialLinkText("PLACE BASKET ORDER")).click();
								w.findElement(By.partialLinkText("Confirm")).click();
															
								String caporderno = w.findElement(By.xpath("//*[@id='container']/div/table[2]/tbody/tr[3]/td[3]")).getText();
																			
								///write square off order no to excel
								j=j+1;
								vr.rw = vr.s.getRow(i);
								vr.cell = vr.rw.createCell(j);
								vr.cell.setCellValue(caporderno);
								
								/////////check for NA order no
								if(caporderno.equals("-NA-"))
								{
									String capmsg = w.findElement(By.xpath("//*[@id='container']/div/table[2]/tbody/tr[3]/td[4]")).getText();
									log.log(LogStatus.INFO, ""+capmsg);
									break;
								}
								else
								{
								log.log(LogStatus.INFO, "Square off order " +caporderno+ " placed successfully");
								w.findElement(By.partialLinkText("Order Status.")).click();
								
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////								
								////modify square off order
								log = report1.startTest("Change " +vr.forreportscripname+ " Square Off Order");
								
								w.findElement(By.xpath(".//input[starts-with(@value,'"+caporderno+"')]")).click();
								log.log(LogStatus.INFO, "Select square off order no "+caporderno);
																	
								String status1 = w.findElement(By.xpath("//td[text()='"+caporderno+"']/following-sibling::td[6]")).getText();
								if(status1.equals("AMO") || status1.equals("OPN"))
								{
												
								///Modify square off order
								w.findElement(By.partialLinkText("CHANGE ORDER")).click();
								log.log(LogStatus.INFO, "Click on Change Order");
								
								/////change order price to actual ltp and trade
								w.findElement(By.id("stk_price")).clear();
								w.findElement(By.id("stk_price")).sendKeys(String.valueOf(vr.pr1));
								log.log(LogStatus.INFO, "Modified at LTP: "+vr.pr1);
								
								w.findElement(By.partialLinkText("CHANGE ORDER")).click();
								w.findElement(By.partialLinkText("Confirm")).click();
								log.log(LogStatus.INFO, "Square off order "+caporderno+ " is modified successfully");
								
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////								
								System.out.println("can new amt "+vr.newamt);
								vr.newamt = vr.cap;
								
								///call capture total
								j=5;
								System.out.println("j new "+j);
								
								capturetotal();
								
								///////to check cancellation util
								Double canamt = vr.newamt - vr.cap;
								System.out.println("can amt " +Math.abs(canamt));
								log.log(LogStatus.INFO, "Cancelled Required Margin is " +Math.abs(canamt));
											
								vr.rw = vr.s.getRow(i);
								vr.cell = vr.rw.createCell(variables.smclutil);
								vr.cell.setCellValue(Math.abs(canamt));
								
								//w.findElement(By.partialLinkText("Order Status")).click();
								Actions act2 = new Actions(w);
								act2.moveToElement(w.findElement(By.linkText("Reports"))).click().perform();	
								WebElement Sublink2 = w.findElement(By.linkText("Order Status"));
								Sublink2.click();
								log.log(LogStatus.INFO, "Click on Reports >> Order Status");
								
								///check SLO after modification
								w.findElement(By.xpath(".//input[starts-with(@value,'"+sno+"')]")).click();
								log.log(LogStatus.INFO, "Select SLO order  "+sno);
								
								String status2 = w.findElement(By.xpath("//td[text()='"+sno+"']/following-sibling::td[6]")).getText();
								if(status2.equals("CAN"))
								{
									log.log(LogStatus.INFO, "Square off order " +caporderno+ " status is " +status2+ " so SLO order got cancelled");
									//break;
								}
								
								else
								{
									log.log(LogStatus.ERROR, "Square off order " +caporderno+ " status is " +status1+ " so SLO order cannot be cancelled");
									//break;
								}
								break;
							  }
								else
								{
									log.log(LogStatus.INFO, "Square off order " +caporderno+ " , SLO order got cancelled");
									
								}
							}
								}
												    	  
				////////////////////////////			
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
			//w.close();
		}
	  
	  @Test(priority=4005)
		public void nouse()
		  {
			  report1.endTest(log);
			  report1.flush();
			  w.get("D:\\testfile\\supermultiple.html");
			  w.get(vr.srcfile);
		  }
	  
}
