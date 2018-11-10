package Testcases;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.List;

import javax.net.ssl.SSLEngineResult.Status;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
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

public class TC6TSLOUtilisation extends BSOrderXL
{
	public static HSSFSheet s;
	public static String popup1,remark,status;
	public static String value;
	public static int i,j,k,rc;
	public static WebElement ele;
	public static String getexactexpdate;
	
	//Utilisation.ManageWatchlist cm = new Utilisation.ManageWatchlist();
	com.library.functions fn = new com.library.functions();
	com.library.variables vr = new com.library.variables();
	
	public static void main(String[] args) 
	  {
		  TestListenerAdapter tla = new TestListenerAdapter();
		    TestNG testng = new TestNG();
		     Class[] classes = new Class[]
		     {
		    	 TC6TSLOUtilisation.class     
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
				vr.cell = vr.rw.createCell(variables.Twputil);
				vr.cell.setCellValue(vr.cap);
			}
			else
			{
				vr.cap = 0.0;
				vr.rw = s.getRow(i);///////////////updated
				vr.cell = vr.rw.createCell(variables.Twputil);///////////////updated
				vr.cell.setCellValue(vr.cap);///////////////updated
				System.out.println("Out of loop");
			}
		}
		else
		{
			
		}	
	}
	
	@Test(priority=0)
	public void PlaceAMOBuyOrder() throws Exception 
	  {
		FileInputStream f = new FileInputStream(srcfile);
		wb = new HSSFWorkbook(f);

		s = wb.getSheet("TSLO");
		
		///report creation
		report1 = new ExtentReports("D:\\Saleema\\TSLOUtilisation.html");
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
			  		vr.forreportscripname=s.getRow(i).getCell(1).getStringCellValue();
				  	System.out.println("reportscripname "+vr.forreportscripname);
				  	log = report1.startTest("Place TSLO " +vr.forreportscripname+ " Order");
	  		
	  				for(j=5;j<row.getLastCellNum();)
	  				{	
					  	///call capture total function to capture utilisation value
	  					//j=5;	
						capturetotal();
						
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
						//wb.close();
						
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
							vr.RM = ((vr.quantity * vr.roundoffpr ) / vr.multiple);
							vr.roundoff = (Math.round(vr.RM*100.0)/100.0);
							log.log(LogStatus.INFO, "Required Margin is " +vr.roundoff);
							System.out.println("roundoff " +vr.roundoff);
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
						vr.rw = s.getRow(i);
						vr.cell = vr.rw.createCell(j);
						vr.cell.setCellValue(vr.pr1);
						
						///write modified price value to excel
						j=j+1;
						vr.rw = s.getRow(i);
						vr.cell = vr.rw.createCell(j);
						vr.cell.setCellValue(vr.roundoffpr);
						
						///chk for TSLO
						j=j+1;
						String tslochk = s.getRow(i).getCell(j).getStringCellValue();
						System.out.println("TSLOchk "+tslochk);
						if(tslochk.equals("YES"))
						{
							//chk checkbox
							w.findElement(By.xpath(".//input[@value= '6']")).click();
							log.log(LogStatus.INFO, "Select TSLO New checkbox");
							
							j=j+1;
							///enter spread value
							float spread = (float) s.getRow(i).getCell(j).getNumericCellValue();
							System.out.println("spread value "+spread);
							w.findElement(By.id("spread")).sendKeys(String.valueOf(spread));
							log.log(LogStatus.INFO, "Spread is: " +spread);
							
							///highlight text
							w.findElement(By.id("slp_amt")).click();
						}
						
						///place order
						w.findElement(By.partialLinkText("PLACE ORDER")).click();
						log.log(LogStatus.INFO, "Click on Place Order button");
						
						///alert
						w.switchTo().alert().accept();
						log.log(LogStatus.INFO, "Click OK on alert");
/////////////////////////////////////////////////////////***********************************************************************						
						///check for any error
						String popup = "Error";
						if (w.getPageSource().contains(popup))
						{	
							popup1 = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
							System.err.println("Error displayed as " +popup1);
							log.log(LogStatus.ERROR, "" +popup1);
							
							vr.rw = s.getRow(i);
							vr.cell = vr.rw.createCell(variables.Twstatus);
							vr.cell.setCellValue(popup1);
							
							if(w.getPageSource().contains("BACK"))
							{
								w.findElement(By.partialLinkText("BACK")).click();
								vr.stckno = null;
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
							String msg = w.findElement(By.xpath("//*[@id='container']/div[1]/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
							System.out.println("msg "+msg);
							log.log(LogStatus.INFO, "" +msg);
							
							///get order no from success message
							vr.stckno = msg.replaceAll("\\D", "");//Integer.parseInt(msg.replaceAll("\\D", ""));
							System.out.println("order no "+vr.stckno);
							log.log(LogStatus.INFO, "Order No is " +vr.stckno);

							w.findElement(By.partialLinkText("SHOW TSLO STATUS")).click();
							
							w.findElement(By.xpath(".//input[@value='"+vr.stckno+",3']")).click();
							
							vr.rw = s.getRow(i);
							vr.cell = vr.rw.createCell(variables.Twpreqmar);
							vr.cell.setCellValue(vr.roundoff);
							
							status = w.findElement(By.xpath("//td[text()='"+vr.stckno+"']/following-sibling::td[10]")).getText();
							Thread.sleep(3000);
							
							System.out.println("status "+status);							
							
							if(!(status.equals("OPN")))//////////////////////////////updated////////////////
							{
								remark = w.findElement(By.xpath("//td[text()='"+vr.stckno+"']/following-sibling::td[11]")).getText();
								log.log(LogStatus.ERROR, "Status is " +status+ " and " +remark+ " so margin utilised cannot be calculated");
							}
							else
							{						
							///click funds >>equity/derivatives
							Actions act1 = new Actions(w);
							act1.moveToElement(w.findElement(By.linkText("Funds"))).click().perform();	
							WebElement Sublink1 = w.findElement(By.linkText("Equity/Derivatives"));
							Sublink1.click();
							
							//click margin utilised cash&FNO
							w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).click();
							
							///check scrip and capture amount
							System.out.println("selected scrip is " +vr.scrip1);
							if(w.getPageSource().contains(vr.scrip1))
							{
								log.log(LogStatus.INFO, "Click on Funds >> Equity/Derivatives");

								///get util amount of specific scrip				
								String total = w.findElement(By.xpath("//td[text()='"+vr.scrip1+"']/following-sibling::td[6]")).getText();
								System.out.println("total " +total);
								vr.tot = Double.parseDouble(total);
								System.out.println("cap amt " +vr.tot);
								Double match = vr.tot - vr.cap;
								System.out.println("match " +match);
								vr.matchround = Math.round(match*100.0)/100.0;
								System.out.println("f amt" +vr.matchround);
								log.log(LogStatus.INFO, "Margin Utilised Cash & FNO " +vr.matchround);
								
								vr.rw = s.getRow(i);
								vr.cell = vr.rw.createCell(variables.Twmutil);
								vr.cell.setCellValue(vr.tot);
								
								vr.rw = s.getRow(i);
								vr.cell = vr.rw.createCell(variables.Twpmarutil);
								vr.cell.setCellValue(vr.matchround);
																
							////to check if reqd margin and captured util amt are same
							if(vr.matchround.equals(vr.roundoff))
							{
								///click on cancel
								w.findElement(By.partialLinkText("Cancel")).click();
								
								///click on total utilised cashNfno
								String totalmargin = w.findElement(By.xpath("//*[@id=\'limit-node-1\']/td[3]")).getText();
								System.out.println(totalmargin);
								String t = totalmargin.replaceAll(",","").trim();
								String m = t.replaceAll("","");
								System.out.println(m);
								log.log(LogStatus.INFO, "Total Margin is " +totalmargin);
								
								vr.rw = s.getRow(i);
								vr.cell = vr.rw.createCell(variables.Twptotmar);
								vr.cell.setCellValue(totalmargin);
								
								///click on margin utilised cashNfno
								String marginutilised = w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).getText();
								System.out.println(marginutilised);
								String t1 = marginutilised.replaceAll(",","").trim();
								String m1 = t1.replaceAll("","");
								System.out.println(m1);
								log.log(LogStatus.INFO, "Margin Utilised is " +marginutilised);
								
								vr.rw = s.getRow(i);
								vr.cell = vr.rw.createCell(variables.Twptotmarutil);
								vr.cell.setCellValue(marginutilised);
								
								///click on margin available cashNfno
								String marginavailable = w.findElement(By.xpath("//*[@id=\'limit-node-6\']/td[3]/strong")).getText();
								System.out.println(marginavailable);
								String t2 = marginavailable.replaceAll(",","").trim();
								String m2 = t2.replaceAll("","");
								System.out.println(m2);
								log.log(LogStatus.INFO, "Margin Available is " +marginavailable);
								
								vr.rw = s.getRow(i);
								vr.cell = vr.rw.createCell(variables.Twpmaravail);
								vr.cell.setCellValue(marginavailable);
								
								//verify totalmargin - margin utilsed = margin available
								Double ab = Double.parseDouble(m)-Double.parseDouble(m1);
								BigDecimal bd = new BigDecimal(ab);
								System.out.println("bd " +bd);
								DecimalFormat df = new DecimalFormat("0.00");
								df.setMaximumFractionDigits(2);
								vr.sd = df.format(ab);
								System.out.println(vr.sd);
								
								vr.rw = s.getRow(i);
								vr.cell = vr.rw.createCell(variables.Twpactmaravail);
								vr.cell.setCellValue(vr.sd);
								
								///check if above are equal
								if(vr.sd.equals(m2))
								{
									System.out.println("Equal");
									log.log(LogStatus.INFO, "Utilized -- As Required Margin and Margin Utilised are Equal");
									log.log(LogStatus.INFO, "Total Margin & Margin utilized are equal to Margin Available " +vr.sd);
																		
									vr.rw = s.getRow(i);
									vr.cell = vr.rw.createCell(variables.Twstatus);
									vr.cell.setCellValue("Pass");
								}
								else
								{
									log.log(LogStatus.ERROR, "Total Margin & Margin utilized are not equal to Margin Available "+vr.sd);
									
									vr.rw = s.getRow(i);
									vr.cell = vr.rw.createCell(variables.Twstatus);
									vr.cell.setCellValue("Fail");
								}
								
							}
							else
							{
								log.log(LogStatus.ERROR, "Not Utilized -- As Required Margin and Margin Utilised are Unequal");
								log.log(LogStatus.ERROR, "Total Margin & Margin utilized are not equal to Margin Available ");
								
								vr.rw = s.getRow(i);
								vr.cell = vr.rw.createCell(variables.Twstatus);
								vr.cell.setCellValue("Fail");
							}
						}
							else
							{
								System.out.println("Selected scrip does not match");
							}
							
						}	
					  }//end of else stckno
						
						///manage order
						log = report1.startTest("Modify TSLO " +vr.forreportscripname+ " Order");
						
						//capturetotal();
						
						///to check for any error
						String amocheck = "Error";
					  	if(w.getPageSource().contains(amocheck))
					  	{
					  		String capturemsg = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
					  		log.log(LogStatus.ERROR, "" +capturemsg);
					  		
					  		vr.rw = s.getRow(i);
					  		vr.cell = vr.rw.createCell(variables.Twstatus);
					  		vr.cell.setCellValue(capturemsg);
					  	}
					  	
					  	else
					  	{
					  	///click watchlist>>place order
						Actions act1 = new Actions(w);
						act1.moveToElement(w.findElement(By.linkText("Reports"))).click().perform();	
						WebElement Sublink1 = w.findElement(By.linkText("TSLO Status"));
						Sublink1.click();
						log.log(LogStatus.INFO, "Click on Reports >> Order Status");
						
					  	//	w.findElement(By.partialLinkText("Order Status")).click();	
					  	
						//check stkno
						System.out.println("manage stk no " +vr.stckno);
						if(vr.stckno !=null &&  !vr.stckno.isEmpty())
						{
						if(w.getPageSource().contains(vr.stckno))
						{
							w.findElement(By.xpath(".//input[@value='"+vr.stckno+",3']")).click();
							log.log(LogStatus.INFO, "Click on TSLO Order No " +vr.stckno);
						}
						Thread.sleep(1000);
						//check status of order no
						String status = w.findElement(By.xpath("//td[text()='"+vr.stckno+"']/following-sibling::td[10]")).getText();
						if((status.equals("OPN")) || status.equals("ACTIVE"))
						{
							
						///click modify order
						w.findElement(By.partialLinkText("MODIFY ORDER")).click();
						
						String check = "BS Indicator";
						if(!w.getPageSource().contains(check))
						{
							//String popup1 = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
							String popup1 = w.findElement(By.xpath("//*[@id='container']/div[1]/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
							log.log(LogStatus.ERROR, "" +popup1);
							//w.findElement(By.partialLinkText("BACK")).click();
							System.out.println("Order no " +vr.stckno+ " cannot be modified as status is " +status);
							
							vr.rw = s.getRow(i);
							vr.cell = vr.rw.createCell(variables.Twstatus);
							vr.cell.setCellValue(popup1);							
						}
						else
						{														
							Double newcap = vr.matchround;
							System.out.println("newcap " +newcap);							
															
							//update stock no
							if(w.findElement(By.id("leg1_ord_qty")).isEnabled())
							{
								/////chk for cash
								j=j+1;
								w.findElement(By.id("leg1_ord_qty")).clear();
								vr.quantity = (int) s.getRow(i).getCell(j).getNumericCellValue();  	
								System.out.println("data1 " +vr.quantity);
								w.findElement(By.id("leg1_ord_qty")).sendKeys(String.valueOf(vr.quantity));
								log.log(LogStatus.INFO, "Modified Quantity is : " +vr.quantity);
								//wb.close();
								
								///modify order
								vr.RM = ((vr.quantity * vr.roundoffpr ) / vr.multiple);
								System.out.println("RM " +vr.RM);
								vr.roundoff1 = (Math.round(vr.RM*100.0)/100.0);
								System.out.println("roundoff " +vr.roundoff1);
							}
							else
							{
								/////chk for fno and currency
								j=j+1;
								w.findElement(By.id("leg1_ord_qty")).clear();
								vr.quantity = (int) s.getRow(i).getCell(j).getNumericCellValue();  	
								w.findElement(By.id("leg1_ord_qty")).sendKeys(String.valueOf(vr.quantity));
								log.log(LogStatus.INFO, "Modified No of lots : " +vr.quantity);
																
								w.findElement(By.id("leg1_ord_prc")).click();
								
								String qt1 = w.findElement(By.id("leg1_ord_qty")).getAttribute("value");
								vr.quantity = Integer.parseInt(qt1);
								log.log(LogStatus.INFO, "Quantity is : " +vr.quantity);
								///modify order
								vr.RM = ((vr.quantity * vr.roundoffpr ) / vr.multiple);
								vr.roundoff1 = (Math.round(vr.RM*100.0)/100.0);
								System.out.println("roundoff " +vr.roundoff1);
							}
							
							//vr.rw = s.getRow(i);
							//vr.cell = vr.rw.createCell(variables.Twmreqmar);
							//vr.cell.setCellValue(vr.roundoff1);
							
							///click on chnge order
							w.findElement(By.partialLinkText("MODIFY")).click();
							log.log(LogStatus.INFO, "Click on Modify Order");	
							
							//cnfrm
							w.findElement(By.partialLinkText("CONFIRM")).click();
							log.log(LogStatus.INFO, "Click Confirm");
							System.out.println("Order no " +vr.stckno+ " is modified successfully");
							log.log(LogStatus.INFO, "Order no " +vr.stckno+ " is modified successfully");
														
							log.log(LogStatus.INFO, "Modified required margin is " +vr.roundoff1);
							
							Actions act15 = new Actions(w);
							act15.moveToElement(w.findElement(By.linkText("Funds"))).click().perform();	
							WebElement Sublink15 = w.findElement(By.linkText("Equity/Derivatives"));
							Sublink15.click();
										
							//capture utilisation total
							w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).click();
							
							///check scrip and capture amount
							
							if(w.getPageSource().contains(vr.scrip1))
							{
								///get util amount of specific scrip				
								String total = w.findElement(By.xpath("//td[text()='"+vr.scrip1+"']/following-sibling::td[6]")).getText();
								System.out.println("mod order total " +total);
								vr.tot = Double.parseDouble(total);
								System.out.println("mod order cap amt " +vr.tot);
								Double match = vr.tot - vr.cap;
								System.out.println("mod order match " +match);
								vr.matchround = Math.round(match*100.0)/100.0;
								System.out.println("mod order mar util matchround" +vr.matchround);
								//log.log(LogStatus.INFO, "Margin Utilised Cash & FNO " +matchround);
								
								
								vr.rw = s.getRow(i);
								vr.cell = vr.rw.createCell(variables.Twcutil);
								vr.cell.setCellValue(vr.tot);
																
																
							////to check if reqd margin and captured util amt are same
							if(vr.matchround.equals(vr.roundoff1))
							{
							///click on cancel
							w.findElement(By.partialLinkText("Cancel")).click();
							
							///click on total utilised cashNfno
							String totalmargin = w.findElement(By.xpath("//*[@id=\'limit-node-1\']/td[3]")).getText();
							System.out.println(totalmargin);
							String t = totalmargin.replaceAll(",","").trim();
							String m = t.replaceAll("","");
							System.out.println("mod order totalmargin " +m);
							log.log(LogStatus.INFO, "Total Margin is " +totalmargin);
							
							vr.rw = s.getRow(i);
							vr.cell = vr.rw.createCell(variables.Twmtotmar);
							vr.rw.removeCell(vr.cell);
							vr.cell.setCellValue(totalmargin);
							
							///click on margin utilised cashNfno
							String marginutilised = w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).getText();
							System.out.println(marginutilised);
							String t1 = marginutilised.replaceAll(",","").trim();
							String m1 = t1.replaceAll("","");
							System.out.println("mod order mar utilised" +m1);
							log.log(LogStatus.INFO, "Margin Utilised is " +marginutilised);
							
							vr.rw = s.getRow(i);
							vr.cell = vr.rw.createCell(variables.Twmtotmarutil);
							vr.rw.removeCell(vr.cell);
							vr.cell.setCellValue(marginutilised);
							
							///click on margin available cashNfno
							String marginavailable = w.findElement(By.xpath("//*[@id=\'limit-node-6\']/td[3]/strong")).getText();
							System.out.println(marginavailable);
							String t2 = marginavailable.replaceAll(",","").trim();
							String m2 = t2.replaceAll("","");
							System.out.println("mod order mar available " +m2);
							log.log(LogStatus.INFO, "Margin Available is " +marginavailable);
							
							vr.rw = s.getRow(i);
							vr.cell = vr.rw.createCell(variables.Twmmaravail);
							vr.rw.removeCell(vr.cell);
							vr.cell.setCellValue(marginavailable);
							
							//verify totalmargin - margin utilsed = margin available
							Double ab = Double.parseDouble(m)-Double.parseDouble(m1);
							BigDecimal bd = new BigDecimal(ab);
							System.out.println("bd " +bd);
							DecimalFormat df = new DecimalFormat("0.00");
							df.setMaximumFractionDigits(2);
							vr.sd = df.format(ab);
							System.out.println("mod order mar util compare " +vr.sd);
							
							vr.rw = s.getRow(i);
							vr.cell = vr.rw.createCell(variables.Twmactmaravail);
							vr.rw.removeCell(vr.cell);
							vr.cell.setCellValue(vr.sd);
							}							
							
						}
							}
						}
						else
						{
							w.findElement(By.partialLinkText("MODIFY ORDER")).click();
							log.log(LogStatus.INFO, "Click on Modify Order");	///////////update here//////////////
							
							String chkerror = "Explanation";
							if(w.getPageSource().contains(chkerror))
							{
								String capturemsg1 = w.findElement(By.xpath("//*[@id='container']/div[1]/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
								log.log(LogStatus.ERROR, "" +remark);
								log.log(LogStatus.ERROR, "" +capturemsg1);
								
								vr.rw = s.getRow(i);
								vr.cell = vr.rw.createCell(variables.Twstatus);
								vr.cell.setCellValue(capturemsg1);
								
								w.findElement(By.partialLinkText("BACK")).click();
								w.findElement(By.xpath(".//input[@value='"+vr.stckno+",3']")).click();
							}
						  }
						}
						else
						{
							System.out.println("Order no not generated as unable to place order.");
							log.log(LogStatus.ERROR, "Order no not generated as " +popup1);
							vr.rw = s.getRow(i);
							vr.cell = vr.rw.createCell(variables.Twstatus);
							vr.cell.setCellValue(popup1);
							
						}
				  	
					  	}	
	  				
	  				
					  	///cancel order
						log = report1.startTest("Cancel TSLO " +vr.forreportscripname+ " Order");
						  
						  ///check for any error
						  String amocheck1 = "Error";
						  if(w.getPageSource().contains(amocheck1))
						  	{
							  	String capturemsg = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
						  		log.log(LogStatus.ERROR, "" +capturemsg);
						  		
						  		vr.rw = s.getRow(i);
						  		vr.cell = vr.rw.createCell(variables.Twstatus);
						  		vr.cell.setCellValue(capturemsg);
						   	}
						  else
						  {
							///click watchlist>>place order
							Actions act2 = new Actions(w);
							act2.moveToElement(w.findElement(By.linkText("Reports"))).click().perform();	
							WebElement Sublink2 = w.findElement(By.linkText("TSLO Status"));
							Sublink2.click();
							log.log(LogStatus.INFO, "Click on Reports >> TSLO Status");
							
							//check stkno
							if(vr.stckno !=null &&  !vr.stckno.isEmpty())
							{
							if(w.getPageSource().contains(vr.stckno))
							{
								w.findElement(By.xpath(".//input[@value='"+vr.stckno+",3']")).click();
								log.log(LogStatus.INFO, "Click on Order No " +vr.stckno);
							}
							
							///click on cancel
							w.findElement(By.partialLinkText("CANCEL ORDER")).click();
							log.log(LogStatus.INFO, "Click on Cancel");
																					
							///check for error
							String trade = "BS Indicator";
							if(!w.getPageSource().contains(trade))
							{
								String popup2 = w.findElement(By.xpath("//*[@id='container']/div[1]/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
								System.out.println("status is "+popup2);
								log.log(LogStatus.ERROR, "" +remark);
								log.log(LogStatus.ERROR, "" +popup2);
								
								vr.rw = s.getRow(i);
								vr.cell = vr.rw.createCell(variables.Twstatus);
								vr.cell.setCellValue(popup2);
								
								w.findElement(By.partialLinkText("BACK")).click();	
								System.out.println("Order no "+vr.stckno+ " cannot be cancelled as status is " +vr.status);
								
								vr.newamt = vr.cap;
								System.out.println("newamt " +vr.newamt);
								
								vr.rw = s.getRow(i);
								vr.cell = vr.rw.createCell(variables.Twclastutil);
								vr.cell.setCellValue(vr.newamt);
							}
							else
							{
								//cnfrm
								w.findElement(By.partialLinkText("CANCEL ORDER")).click();
								log.log(LogStatus.INFO, "Click on Cancel Order");
								System.out.println("Order no "+vr.stckno+ " is cancelled successfully");
								log.log(LogStatus.INFO, "Order no " +vr.stckno+ " is cancelled successfully");
									
								if(status.equals("OPN"))
								{
									vr.newamt = vr.cap;
									
									vr.rw = s.getRow(i);
									vr.cell = vr.rw.createCell(variables.Twclastutil);
									vr.cell.setCellValue(vr.newamt);
									
									///call capture total
									j=5;
									System.out.println("j new "+j);
									capturetotal();
									Double canamt = vr.newamt - vr.cap;
									System.out.println("can amt " +Math.abs(canamt));
									log.log(LogStatus.INFO, "Cancelled Required Margin is " +Math.abs(canamt));
																	
									///click on cancel
									w.findElement(By.partialLinkText("Cancel")).click();
									
									///click on total utilised cashNfno
									String totalmargin = w.findElement(By.xpath("//*[@id=\'limit-node-1\']/td[3]")).getText();
									System.out.println(totalmargin);
									String t = totalmargin.replaceAll(",","").trim();
									String m = t.replaceAll("","");
									System.out.println("total can order " +m);
									log.log(LogStatus.INFO, "Total Margin is " +totalmargin);
									
									vr.rw = s.getRow(i);
									vr.cell = vr.rw.createCell(variables.Twctotmar);
									vr.cell.setCellValue(totalmargin);
									
									///click on margin utilised cashNfno
									String marginutilised = w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).getText();
									System.out.println(marginutilised);
									String t1 = marginutilised.replaceAll(",","").trim();
									String m1 = t1.replaceAll("","");
									System.out.println("can order mar util " +m1);
									log.log(LogStatus.INFO, "Margin Utilised is " +marginutilised);
									
									vr.rw = s.getRow(i);
									vr.cell = vr.rw.createCell(variables.Twctotmarutil);
									vr.cell.setCellValue(marginutilised);
									
									///click on margin available cashNfno
									String marginavailable = w.findElement(By.xpath("//*[@id=\'limit-node-6\']/td[3]/strong")).getText();
									System.out.println(marginavailable);
									String t2 = marginavailable.replaceAll(",","").trim();
									String m2 = t2.replaceAll("","");
									System.out.println("can order mar avail " +m2);
									log.log(LogStatus.INFO, "Margin Available is " +marginavailable);
									
									vr.rw = s.getRow(i);
									vr.cell = vr.rw.createCell(variables.Twcmaravail);
									vr.cell.setCellValue(marginavailable);
									
									//verify totalmargin - margin utilsed = margin available
									Double ab = Double.parseDouble(m)-Double.parseDouble(m1);
									BigDecimal bd = new BigDecimal(ab);
									System.out.println("bd " +bd);
									DecimalFormat df = new DecimalFormat("0.00");
									df.setMaximumFractionDigits(2);
									vr.sd = df.format(ab);
									System.out.println("can order mar util capmare " +vr.sd);
									
									vr.rw = s.getRow(i);
									vr.cell = vr.rw.createCell(variables.Twcactmaravail);
									vr.cell.setCellValue(vr.sd);
							}
							}
						  
						  }
							else
							{
								System.out.println("Order no not generated as unable to place order.");
								log.log(LogStatus.ERROR, "Order no not generated as " +popup1);
								vr.rw = s.getRow(i);
								vr.cell = vr.rw.createCell(variables.Twstatus);
								vr.cell.setCellValue(popup1);
							}
						  }
						k=j;
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
	 
	@AfterTest
	public void aftertest()
	{
		//w.close();
	}
	
	@Test(priority=5001)
	public void nouse()
	  {
		  report1.endTest(log);
		  report1.flush();
		  w.get("D:\\Saleema\\TSLOUtilisation.html");
		  //w.get(srcfile);
	  }			  
	  
}
