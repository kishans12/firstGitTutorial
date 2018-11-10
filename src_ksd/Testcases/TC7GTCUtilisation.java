package Testcases;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import org.apache.commons.lang3.text.WordUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
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

public class TC7GTCUtilisation extends BSOrderXL
{
	public static String gtcorderno,chkitms,sd;
	public static HSSFSheet s;
	public static Double roundoff;
	public static Double RM;
	public static int multiple;
	public static Double price;
	public static Double matchround;
	public static Double tot;
	
	public static String popup1,value,getexactexpdate;
	public static int i,j,k,rc;
	public static WebElement ele;
	
	com.library.functions fn = new com.library.functions();
	com.library.variables vr = new com.library.variables();			
	
	public static void main(String[] args) 
	  {
		  TestListenerAdapter tla = new TestListenerAdapter();
		    TestNG testng = new TestNG();
		     Class[] classes = new Class[]
		     {
		    	 TC7GTCUtilisation.class     
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
	
	public void date() throws Exception
	{
		String window = w.getWindowHandle();
		w.findElement(By.xpath("//*[@id=\'gtctable\']/tbody/tr[2]/td[8]/a/img")).click();
		System.out.println(w.getTitle());
		for(String winHandle : w.getWindowHandles())
		{
		    w.switchTo().window(winHandle);
		}
		System.out.println(w.getTitle());
		
		//to print month
		Date today = Calendar.getInstance().getTime();
		SimpleDateFormat formatter = new SimpleDateFormat("MM");
		String month1 = formatter.format(today);
		
		//to print year
		SimpleDateFormat year = new SimpleDateFormat("yyyy");
		String yr = year.format(today);
		
		//to print day
		DateFormat dateFormat = new SimpleDateFormat("d");
		Calendar cal = Calendar.getInstance();
		cal.setTime(new Date());
		cal.add(Calendar.DATE, 1);
		String newDate = dateFormat.format(cal.getTime());
		System.out.println("newdate "+newDate);
		//click on next day
		w.findElement(By.xpath("//a[text()= '"+newDate+"']")).click();
		w.switchTo().window(window);
		System.out.println("back to window");
		log.log(LogStatus.INFO, "Selected Date is "+newDate+ "/" +month1+ "/" +yr);
	}
	
	@Test(priority=0)
	public void PlaceGTCOrder() throws Exception
	{
		FileInputStream f = new FileInputStream(srcfile);
		HSSFWorkbook wb = new HSSFWorkbook(f);
		
		s = wb.getSheet("GTC");
		
		///report creation
		report1 = new ExtentReports("D:\\testfile\\GTCUtilisation.html");
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
				  	log = report1.startTest("Place GTC " +vr.forreportscripname+ " Order");
				  	
	  				for(j=5;j<row.getLastCellNum();)
	  				{
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
						
						///click GTC
						w.findElement(By.partialLinkText("GTC ORDER")).click();
						log.log(LogStatus.INFO, "Click on GTC Order");											
						
						///get scripname
						String scripname = w.findElement(By.id("name1")).getAttribute("value");
						log.log(LogStatus.INFO, "Selected Scrip is " +scripname);
						
						///Select buy/sell value
						j=j+1;
						w.findElement(By.id("bsind1")).click();
						value = s.getRow(i).getCell(j).getStringCellValue();
						Select val = new Select(w.findElement(By.id("bsind1")));
						val.selectByVisibleText(org.apache.commons.text.WordUtils.capitalizeFully(value));
						log.log(LogStatus.INFO, "Selected Value is " +value);
						
						////enter quantity
						j=j+1;
						vr.quantity = (int) s.getRow(i).getCell(j).getNumericCellValue();  	
						w.findElement(By.id("qty1")).sendKeys(String.valueOf(vr.quantity));
						log.log(LogStatus.INFO, "Entered Quantity is: " +vr.quantity);
						
						///multiple
						j=j+1;
						vr.multiple = (float) s.getRow(i).getCell(j).getNumericCellValue();  	
						System.out.println("multiple " +vr.multiple);
						
						///required margin calculation
						j=j+1;
						int perc = (int) s.getRow(i).getCell(j).getNumericCellValue();
						vr.pr1 = w.findElement(By.id("price1")).getAttribute("value");
						w.findElement(By.id("price1")).clear();
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
						w.findElement(By.id("price1")).sendKeys(String.valueOf(vr.roundoffpr));
						log.log(LogStatus.INFO, "Actual price is " +vr.pr1);
						log.log(LogStatus.INFO, "Modified price is " +vr.roundoffpr);
						vr.RM = ((vr.quantity * vr.roundoffpr ) / vr.multiple);
						vr.roundoff = (Math.round(vr.RM*100.0)/100.0);
						log.log(LogStatus.INFO, "Required Margin is " +vr.roundoff);
						System.out.println("roundoff " +vr.roundoff);
						
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
						
						///select date
						date();
						
						///click on place order
						w.findElement(By.xpath("//*[@id='container']/div/form/table/tbody/tr/td[1]/a")).click();
						log.log(LogStatus.INFO, "Click on Place Order");
						
						String error = "Error";
						if(w.getPageSource().contains(error))
						{
							String capturemsg = w.findElement(By.xpath("//*[@id='container']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
							System.err.println("Error displayed as " +capturemsg);
							log.log(LogStatus.ERROR, "" +capturemsg);
						}
						else
						{	
							///click on confirm
							w.findElement(By.partialLinkText("Confirm")).click();
							log.log(LogStatus.INFO, "Click on Confirm");
							
							////capture gtc order no
							gtcorderno = w.findElement(By.xpath("//*[@id='container']/div/form/table[3]/tbody/tr[3]/td[3]")).getText();
							log.log(LogStatus.INFO, "GTC Order got placed successfully and Order No is " +gtcorderno);
							
							vr.rw = s.getRow(i);
							vr.cell = vr.rw.createCell(variables.wpreqmar);
							vr.cell.setCellValue(vr.roundoff);
							
////////////////////////////////////////////check if order is active/open --- place//////////////////////////////////////////////
							//Thread.sleep(20000);		/////for placing order from dealer
							///click watchlist>>place order
							Actions act3 = new Actions(w);
							act3.moveToElement(w.findElement(By.linkText("Reports"))).click().perform();	
							WebElement Sublink3 = w.findElement(By.linkText("GTC Status"));
							Sublink3.click();
							log.log(LogStatus.INFO, "Click on Reports >> GTC Status");
							
							chkitms = w.findElement(By.xpath("//input[starts-with(@value,'"+gtcorderno+"')]/following-sibling::td[2]")).getText();
							System.out.println("chkitms manage "+chkitms);
							
							///chk if gtc order is active
							if(chkitms.equals("0"))
							{
								System.out.println("gtc order is active");
								log.log(LogStatus.INFO, "GTC Order " +gtcorderno+ " is Active so utilisation cannot be calculated");
								
								vr.rw = s.getRow(i);
								vr.cell = vr.rw.createCell(variables.wstatus);
								vr.cell.setCellValue("Status is ACTIVE so utilisation cannot be calculated");
							}
							else
							{							
								/////if gtc order is open
																
								///click watchlist>>place order
								Actions act11 = new Actions(w);
								act11.moveToElement(w.findElement(By.linkText("Reports"))).click().perform();	
								WebElement Sublink11 = w.findElement(By.linkText("Order Status"));
								Sublink11.click();
								log.log(LogStatus.INFO, "Click on Reports >> Order Status");
								
								if(chkitms !=null &&  !chkitms.isEmpty())
								{
									if(w.getPageSource().contains(chkitms))
									{
										w.findElement(By.xpath(".//input[starts-with(@value,'"+chkitms+"')]")).click();
										log.log(LogStatus.INFO, "Click on Order No " +chkitms);
									}
				
									//check status of order no
									String status = w.findElement(By.xpath("//td[text()='"+chkitms+"']/following-sibling::td[6]")).getText();
									if(status.equals("OPN"))
									{
										///click funds >>equity/derivatives
										Actions act1 = new Actions(w);
										act1.moveToElement(w.findElement(By.linkText("Funds"))).click().perform();	
										WebElement Sublink1 = w.findElement(By.linkText("Equity/Derivatives"));
										Sublink1.click();
										log.log(LogStatus.INFO, "Click on Funds >> Equity/Derivatives");
										
										//click margin utilised cash&FNO
										w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).click();
										///check scrip and capture amount
										System.out.println("selected scrip is " +vr.scrip1);
										if(w.getPageSource().contains(vr.scrip1))
										{
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
											vr.cell = vr.rw.createCell(variables.wmutil);
											vr.cell.setCellValue(vr.tot);
											
											vr.rw = s.getRow(i);
											vr.cell = vr.rw.createCell(variables.wpmarutil);
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
											vr.cell = vr.rw.createCell(variables.wptotmar);
											vr.cell.setCellValue(totalmargin);
											
											///click on margin utilised cashNfno
											String marginutilised = w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).getText();
											System.out.println(marginutilised);
											String t1 = marginutilised.replaceAll(",","").trim();
											String m1 = t1.replaceAll("","");
											System.out.println(m1);
											log.log(LogStatus.INFO, "Margin Utilised is " +marginutilised);
											
											vr.rw = s.getRow(i);
											vr.cell = vr.rw.createCell(variables.wptotmarutil);
											vr.cell.setCellValue(marginutilised);
											
											///click on margin available cashNfno
											String marginavailable = w.findElement(By.xpath("//*[@id=\'limit-node-6\']/td[3]/strong")).getText();
											System.out.println(marginavailable);
											String t2 = marginavailable.replaceAll(",","").trim();
											String m2 = t2.replaceAll("","");
											System.out.println(m2);
											log.log(LogStatus.INFO, "Margin Available is " +marginavailable);
											
											vr.rw = s.getRow(i);
											vr.cell = vr.rw.createCell(variables.wpmaravail);
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
											vr.cell = vr.rw.createCell(variables.wpactmaravail);
											vr.cell.setCellValue(vr.sd);
											
											///check if above are equal
											if(vr.sd.equals(m2))
											{
												System.out.println("Equal");
												log.log(LogStatus.INFO, "Utilized -- As Required Margin and Margin Utilised are Equal");
												log.log(LogStatus.INFO, "Total Margin & Margin utilized are equal to Margin Available " +vr.sd);
																					
												vr.rw = s.getRow(i);
												vr.cell = vr.rw.createCell(variables.wstatus);
												vr.cell.setCellValue("Pass");
											}
											else
											{
												log.log(LogStatus.ERROR, "Total Margin & Margin utilized are not equal to Margin Available "+vr.sd);
												
												vr.rw = s.getRow(i);
												vr.cell = vr.rw.createCell(variables.wstatus);
												vr.cell.setCellValue("Fail");
											}
											
										}
										else
										{
											log.log(LogStatus.ERROR, "Not Utilized -- As Required Margin and Margin Utilised are Unequal");
											log.log(LogStatus.ERROR, "Total Margin & Margin utilized are not equal to Margin Available ");
											
											vr.rw = s.getRow(i);
											vr.cell = vr.rw.createCell(variables.wstatus);
											vr.cell.setCellValue("Fail");
										}
									}
										else
										{
											System.out.println("Selected scrip does not match");
										}
									}
									else
									{
										System.out.println("gtc order is not open ");
									}
								}
								else
								{
									System.out.println("Order no not generated as unable to place order.");
									log.log(LogStatus.ERROR, "Order no not generated as " +popup1);									
								}																							
								
							}
							
						}
						
						///manage gtc order
						log = report1.startTest("Change GTC " +vr.forreportscripname+ " Order");
						
						///to check for any error
						String amocheck = "Error";
					  	if(w.getPageSource().contains(amocheck))
					  	{
					  		String capturemsg = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
					  		log.log(LogStatus.ERROR, "" +capturemsg);
					  		
					  		vr.rw = s.getRow(i);
					  		vr.cell = vr.rw.createCell(variables.wstatus);
					  		vr.cell.setCellValue(capturemsg);
					  	}
					  	
					  	else
					  	{
					  		///click watchlist>>place order
							Actions act3 = new Actions(w);
							act3.moveToElement(w.findElement(By.linkText("Reports"))).click().perform();	
							WebElement Sublink3 = w.findElement(By.linkText("GTC Status"));
							Sublink3.click();
							log.log(LogStatus.INFO, "Click on Reports >> GTC Status");
							
							chkitms = w.findElement(By.xpath("//input[starts-with(@value,'"+gtcorderno+"')]/following-sibling::td[2]")).getText();
							System.out.println("chkitms manage "+chkitms);
							
							log.log(LogStatus.INFO, "Select GTC order No " +gtcorderno);

							///chk if gtc order is active
							if(chkitms.equals("0"))
							{								
								///click on modify
								w.findElement(By.xpath("//input[starts-with(@value,'"+gtcorderno+"')]/following-sibling::td[12]/a/img")).click();
								log.log(LogStatus.INFO, "Click on Modify");
								
								String error1 = "Error";
								if(w.getPageSource().contains(error1))
								{
									String capturemsg = w.findElement(By.xpath("//*[@id=\'container\']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
									System.err.println("Error displayed as " +capturemsg);
									log.log(LogStatus.ERROR, "" +capturemsg);
								}
								
								Double newcap = vr.matchround;
								System.out.println("newcap " +newcap);
								
								///modify quantity
								j=j+2;
								w.findElement(By.id("qty")).clear();
								vr.quantity = (int) s.getRow(i).getCell(j).getNumericCellValue();  	
								w.findElement(By.id("qty")).sendKeys(String.valueOf(vr.quantity));
								log.log(LogStatus.INFO, "Modified Quantity : " +vr.quantity);
								
								///modify order
								vr.RM = ((vr.quantity * vr.roundoffpr ) / vr.multiple);
								System.out.println("RM " +vr.RM);
								vr.roundoff1 = (Math.round(vr.RM*100.0)/100.0);
								System.out.println("roundoff " +vr.roundoff1);
								
								//vr.rw = s.getRow(i);
								//vr.cell = vr.rw.createCell(variables.wmreqmar);
								//vr.cell.setCellValue(vr.roundoff1);
								
								//click place order
								w.findElement(By.xpath("//*[@id=\'container\']/div/form/table[3]/tbody/tr/td[1]/a")).click();
								log.log(LogStatus.INFO, "Click on Place Order");
								
								//confirm
								w.findElement(By.partialLinkText("CONFIRM")).click();
								log.log(LogStatus.INFO, "Click on Confirm");
								
								log.log(LogStatus.INFO, "Order No " +gtcorderno+ " is modified successfully");
						}
							else
							{
								////itms not 0 else
								log.log(LogStatus.INFO, "GTC Order " +gtcorderno+ " is not Active so order cannot be modified");
							}							
					  	}
					  	
					  	///cancel order
						log = report1.startTest("Cancel GTC " +vr.forreportscripname+ " Order");
						
						///check for any error
						String cancheck = "Error";
						if(w.getPageSource().contains(cancheck))
						{
							String capturemsg = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
						  	log.log(LogStatus.ERROR, "" +capturemsg);
						  	
						  	vr.rw = s.getRow(i);
					  		vr.cell = vr.rw.createCell(variables.wstatus);
					  		vr.cell.setCellValue(capturemsg);						  	
						}
						  else
						  {
							  	///select status
								//w.findElement(By.partialLinkText("Status")).click();
								//log.log(LogStatus.INFO, "Click on Status");
							  	
							  	///click watchlist>>place order
								Actions act4 = new Actions(w);
								act4.moveToElement(w.findElement(By.linkText("Reports"))).click().perform();	
								WebElement Sublink4 = w.findElement(By.linkText("GTC Status"));
								Sublink4.click();
								log.log(LogStatus.INFO, "Click on Reports >> GTC Status");
								
								log.log(LogStatus.INFO, "Select GTC order No " +gtcorderno);
								
								if(chkitms.equals("0"))
								{
									//click cancel
									w.findElement(By.xpath("//input[starts-with(@value,'"+gtcorderno+"')]/following-sibling::td[13]/a/img")).click();
									log.log(LogStatus.INFO, "Click on Cancel");
								}
								else
								{
									//click cancel
									w.findElement(By.xpath("//input[starts-with(@value,'"+gtcorderno+"')]/following-sibling::td[12]/a/img")).click();
									log.log(LogStatus.INFO, "Click on Cancel");
								}								
								
								String err = "Error";
								if(w.getPageSource().contains(err))
								{
									String capturemsg = w.findElement(By.xpath("//*[@id=\'container\']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
									log.log(LogStatus.ERROR, "" +capturemsg);
									
									vr.rw = s.getRow(i);
									vr.cell = vr.rw.createCell(variables.wstatus);
									vr.cell.setCellValue(capturemsg);
									
									w.findElement(By.partialLinkText("BACK")).click();
									vr.newamt = vr.cap;
									System.out.println("newamt " +vr.newamt);
									
									vr.rw = s.getRow(i);
									vr.cell = vr.rw.createCell(variables.wclastutil);
									vr.cell.setCellValue(vr.newamt);
								}
								else
								{
									//click confirm
									w.findElement(By.partialLinkText("Confirm")).click();
									log.log(LogStatus.INFO, "Click on Confirm");									
									log.log(LogStatus.INFO, "Order No " +gtcorderno+ " is cancelled successfully");
									
									if(!chkitms.equals("0"))
									{											
										vr.newamt = vr.cap;
										vr.rw = s.getRow(i);
										vr.cell = vr.rw.createCell(variables.wclastutil);
										vr.cell.setCellValue(vr.newamt);
										
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
										vr.cell = vr.rw.createCell(variables.wctotmar);
										vr.cell.setCellValue(totalmargin);
										
										///click on margin utilised cashNfno
										String marginutilised = w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).getText();
										System.out.println(marginutilised);
										String t1 = marginutilised.replaceAll(",","").trim();
										String m1 = t1.replaceAll("","");
										System.out.println("can order mar util " +m1);
										log.log(LogStatus.INFO, "Margin Utilised is " +marginutilised);
										
										vr.rw = s.getRow(i);
										vr.cell = vr.rw.createCell(variables.wctotmarutil);
										vr.cell.setCellValue(marginutilised);
										
										///click on margin available cashNfno
										String marginavailable = w.findElement(By.xpath("//*[@id=\'limit-node-6\']/td[3]/strong")).getText();
										System.out.println(marginavailable);
										String t2 = marginavailable.replaceAll(",","").trim();
										String m2 = t2.replaceAll("","");
										System.out.println("can order mar avail " +m2);
										log.log(LogStatus.INFO, "Margin Available is " +marginavailable);
										
										vr.rw = s.getRow(i);
										vr.cell = vr.rw.createCell(variables.wcmaravail);
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
										vr.cell = vr.rw.createCell(variables.wcactmaravail);
										vr.cell.setCellValue(vr.sd);
									}
								}
						  }
						k=j;
						break;
						
	  				}//j loop
	  				
				}//flag loop
				
				else
	  			{
	  				System.out.println("Flag is N");
	  			}
				
	  	}//i loop
	  	
	  	FileOutputStream Of = new FileOutputStream(srcfile);
	  	
		wb.write(Of);
		Of.close();
							
	}///end loop
	
	
	@AfterTest
	public void aftertest()
	{
		//w.close();
	}	
		
	@Test(priority=3634)
	public void nouse()
	  {
		 report1.endTest(log);
		 report1.flush();
		 w.get("D:\\testfile\\GTCUtilisation.html");
		 w.get(srcfile);
	  }
}
