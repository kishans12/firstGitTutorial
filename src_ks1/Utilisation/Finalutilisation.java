package Utilisation;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;

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

public class Finalutilisation extends BSOrderXL
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
	public static String value;
	public String stckno;
	public String Status1;
	public String status;
	public static int j;
	public static int i;
	public static int k;
	public static Row rw;
	public static Cell cell;
			
	Utilisation.ManageWatchlist cm = new Utilisation.ManageWatchlist();
	//com.library.functions fn = new com.library.functions();
	//com.library.variables vb = new com.library.variables();
	
	public void capturetotal() throws Exception
	{
		HSSFSheet s = wb.getSheet("Data");
		HSSFSheet sw = wb.getSheet("Result");
				
		///click funds >>equity/derivatives
		Actions act1 = new Actions(w);
		act1.moveToElement(w.findElement(By.linkText("Funds"))).click().perform();	
		WebElement Sublink1 = w.findElement(By.linkText("Equity/Derivatives"));
		Sublink1.click();
					
		//capture utilisation total
		w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).click();
		String checkmsg = "No Record Found";
				
		scrip1 = s.getRow(i).getCell(j).getStringCellValue();
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
			rw = sw.getRow(i);
			cell = rw.createCell(variables.wputil);
			cell.setCellValue(cap);
		}
		else
		{
			System.out.println("Out of loop");
		}
				
	}
	
	@Test(priority=0)
	public void PlaceAMOBuyOrder() throws Exception 
	  {
		//FileInputStream f = new FileInputStream(srcfile);
		//wb = new HSSFWorkbook(f);
	  	FileOutputStream Of = new FileOutputStream(srcfile);

		HSSFSheet s = wb.getSheet("Data");
		HSSFSheet s1 = wb.getSheet("AddScrip");
		HSSFSheet sw = wb.getSheet("Result");
		
		///report creation
		report1 = new ExtentReports("D:\\Saleema\\looputil.html");
		log = report1.startTest("Manage Watchlist");
	  	
		////call login function
	  	Login();
	  	
	  	///call manage watchlist
	  	cm.manage();
	  		  	
	  	///get multiple value from excel for utilization caln
	  	
	  	int rc = s.getLastRowNum() - s.getFirstRowNum();
	  	System.out.println("rc "+rc);
	  	String watchlistname = s1.getRow(1).getCell(1).getStringCellValue();
	  	
	  	for(i=5;i<rc+1;i++)
	  	{
	  		Row row = s.getRow(i);
	  		
	  		String forreportscripname=s.getRow(i).getCell(1).getStringCellValue();
		  	System.out.println("reportscripname "+forreportscripname);
		  	log = report1.startTest("Place " +forreportscripname+ " Order");
	  		
	  				for(j=5;j<row.getLastCellNum();)
	  				{	  					  					
	  					String flag = s.getRow(i).getCell(j).getStringCellValue();
	  					
	  					if(flag.equals("Y"))
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
						String xcelscrip = s.getRow(i).getCell(j).getStringCellValue();
						System.out.println("data1 " +xcelscrip);
						w.findElement(By.xpath(".//input[starts-with(@value,'"+xcelscrip+"')]")).click();
						
						j=j+1;
						String getexccode = w.findElement(By.xpath(".//td[input[starts-with(@value,'"+xcelscrip+"')]]/following-sibling::td[3]")).getText();
						System.out.println("getexchange code "+getexccode);
						
						rw = s.getRow(i);
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
						value = s.getRow(i).getCell(j).getStringCellValue();
						System.out.println("boolean " +value);
						w.findElement(By.xpath(".//input[@value='"+value+"']")).click();
						log.log(LogStatus.INFO, "Selected Value is " +value);
											
						////enter quantity
						j=j+1;
						int quantity = (int) s.getRow(i).getCell(j).getNumericCellValue();  	
						System.out.println("data1 " +quantity);
						//wb.close();
						
						if(w.findElement(By.id("stk_qty")).isEnabled())
						{
							w.findElement(By.id("stk_qty")).sendKeys(String.valueOf(quantity));
							log.log(LogStatus.INFO, "Entered Quantity is: " +quantity);
							
							///multiple
							j=j+1;
							multiple = (float) s.getRow(i).getCell(j).getNumericCellValue();  	
							System.out.println("multiple " +multiple);
							
							///required margin calculation
							j=j+1;
							int perc = (int) s.getRow(i).getCell(j).getNumericCellValue();
							String pr1 = w.findElement(By.id("stk_price")).getAttribute("value");
							w.findElement(By.id("stk_price")).clear();
							price = (Double.parseDouble(pr1)*perc)/100;
							double chngestrngpr = Double.parseDouble(pr1);
							if(value.equals("BUY"))
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
							multiple = (float) s.getRow(i).getCell(j).getNumericCellValue();  	
							System.out.println("multiple  " +multiple);
							
							///required margin calculation
							j=j+1;
							int perc = (int) s.getRow(i).getCell(j).getNumericCellValue();
							String pr1 = w.findElement(By.id("stk_price")).getAttribute("value");
							w.findElement(By.id("stk_price")).clear();
							price = (Double.parseDouble(pr1)*perc)/100;
							double chngestrngpr = Double.parseDouble(pr1);
							if(value.equals("BUY"))
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
						
						///write modified price value to excel
						j=j+1;
						rw = s.getRow(i);
						cell = rw.createCell(j);
						cell.setCellValue(roundoffpr);
						
						///chk for AMO
						j=j+1;
						String amochk = s.getRow(i).getCell(j).getStringCellValue();
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
							stckno = w.findElement(By.xpath("//*[@id='container']/div/p/span")).getText();
							System.out.println("stckno...... " +stckno);
							log.log(LogStatus.INFO, "Order No is " +stckno);
							
							rw = sw.getRow(i);
							cell = rw.createCell(variables.wpreqmar);
							cell.setCellValue(roundoff);
							
							///click funds >>equity/derivatives
							Actions act1 = new Actions(w);
							act1.moveToElement(w.findElement(By.linkText("Funds"))).click().perform();	
							WebElement Sublink1 = w.findElement(By.linkText("Equity/Derivatives"));
							Sublink1.click();
							log.log(LogStatus.INFO, "Click on Funds >> Equity/Derivatives");
							
							//click margin utilised cash&FNO
							w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).click();
							
							///check scrip and capture amount
							
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
								log.log(LogStatus.INFO, "Margin Utilised Cash & FNO " +matchround);
								
								rw = sw.getRow(i);
								cell = rw.createCell(variables.wmutil);
								cell.setCellValue(tot);
								
								rw = sw.getRow(i);
								cell = rw.createCell(variables.wpmarutil);
								cell.setCellValue(matchround);
																
							////to check if reqd margin and captured util amt are same
							if(matchround.equals(roundoff))
							{
								///click on cancel
								w.findElement(By.partialLinkText("CANCEL")).click();
								
								///click on total utilised cashNfno
								String totalmargin = w.findElement(By.xpath("//*[@id=\'limit-node-1\']/td[3]")).getText();
								System.out.println(totalmargin);
								String t = totalmargin.replaceAll(",","").trim();
								String m = t.replaceAll("","");
								System.out.println(m);
								log.log(LogStatus.INFO, "Total Margin is " +totalmargin);
								
								rw = sw.getRow(i);
								cell = rw.createCell(variables.wptotmar);
								rw.removeCell(cell);
								cell.setCellValue(totalmargin);
								
								///click on margin utilised cashNfno
								String marginutilised = w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).getText();
								System.out.println(marginutilised);
								String t1 = marginutilised.replaceAll(",","").trim();
								String m1 = t1.replaceAll("","");
								System.out.println(m1);
								log.log(LogStatus.INFO, "Margin Utilised is " +marginutilised);
								
								rw = sw.getRow(i);
								cell = rw.createCell(variables.wptotmarutil);
								rw.removeCell(cell);
								cell.setCellValue(marginutilised);
								
								///click on margin available cashNfno
								String marginavailable = w.findElement(By.xpath("//*[@id=\'limit-node-6\']/td[3]/strong")).getText();
								System.out.println(marginavailable);
								String t2 = marginavailable.replaceAll(",","").trim();
								String m2 = t2.replaceAll("","");
								System.out.println(m2);
								log.log(LogStatus.INFO, "Margin Available is " +marginavailable);
								
								rw = sw.getRow(i);
								cell = rw.createCell(variables.wpmaravail);
								rw.removeCell(cell);
								cell.setCellValue(marginavailable);
								
								//verify totalmargin - margin utilsed = margin available
								Double ab = Double.parseDouble(m)-Double.parseDouble(m1);
								BigDecimal bd = new BigDecimal(ab);
								System.out.println("bd " +bd);
								DecimalFormat df = new DecimalFormat("0.00");
								df.setMaximumFractionDigits(2);
								sd = df.format(ab);
								System.out.println(sd);
								
								rw = sw.getRow(i);
								cell = rw.createCell(variables.wpactmaravail);
								rw.removeCell(cell);
								cell.setCellValue(sd);
								
								///check if above are equal
								if(sd.equals(m2))
								{
									System.out.println("Equal");
									log.log(LogStatus.INFO, "Utilized -- As Required Margin and Margin Utilised are Equal");
									log.log(LogStatus.INFO, "Total Margin & Margin utilized are equal to Margin Available " +sd);
																		
									rw = sw.getRow(i);
									cell = rw.createCell(variables.wstatus);
									rw.removeCell(cell);
									cell.setCellValue("Pass");
								}
								else
								{
									log.log(LogStatus.ERROR, "Total Margin & Margin utilized are not equal to Margin Available "+sd);
									
									rw = sw.getRow(i);
									cell = rw.createCell(variables.wstatus);
									rw.removeCell(cell);
									cell.setCellValue("Fail");
								}
								
							}
							else
							{
								log.log(LogStatus.ERROR, "Not Utilized -- As Required Margin and Margin Utilised are Unequal");
								log.log(LogStatus.ERROR, "Total Margin & Margin utilized are not equal to Margin Available ");
								
								rw = sw.getRow(i);
								cell = rw.createCell(variables.wstatus);
								cell.setCellValue("Fail");
							}
						}
							else
							{
								System.out.println("Selected scrip does not match");
							}
							
							//FileOutputStream Of = new FileOutputStream(srcfile);
							//wb.write(Of);
							//Of.close();
					  }
						
						///manage order
						log = report1.startTest("Change " +forreportscripname+ " Order");
						
						//capturetotal();
						
						///to check for any error
						String amocheck = "Error";
					  	if(w.getPageSource().contains(amocheck))
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
						
					  	//	w.findElement(By.partialLinkText("Order Status")).click();	
					  	
						//check stkno
						System.out.println("manage stk no " +stckno);
						if(stckno !=null &&  !stckno.isEmpty())
						{
						if(w.getPageSource().contains(stckno))
						{
						w.findElement(By.xpath(".//input[starts-with(@value,'"+stckno+"')]")).click();
						log.log(LogStatus.INFO, "Click on Order No " +stckno);
						}
						Thread.sleep(1000);
						//check status of order no
						String status = w.findElement(By.xpath("//td[text()='"+stckno+"']/following-sibling::td[6]")).getText();
						if(status.equals("AMO") || status.equals("OPN"))
						{
										
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
							j=j+1;
							w.findElement(By.id("stk_qty")).clear();
							quantity = (int) s.getRow(i).getCell(j).getNumericCellValue();  	
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
								j=j+1;
								w.findElement(By.id("stk_lot")).clear();
								quantity = (int) s.getRow(i).getCell(j).getNumericCellValue();  	
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
							
							rw = sw.getRow(i);
							cell = rw.createCell(variables.wmreqmar);
							rw.removeCell(cell);
							cell.setCellValue(roundoff1);
							
							///click on chnge order
							w.findElement(By.partialLinkText("CHANGE ORDER")).click();
							log.log(LogStatus.INFO, "Click on Change Order");	
							
							//cnfrm
							w.findElement(By.partialLinkText("Confirm")).click();
							log.log(LogStatus.INFO, "Click Confirm");
							System.out.println("Order no " +stckno+ " is modified successfully");
							log.log(LogStatus.INFO, "Order no " +stckno+ " is modified successfully");
														
							log.log(LogStatus.INFO, "Modified required margin is " +roundoff1);
							
							Actions act15 = new Actions(w);
							act15.moveToElement(w.findElement(By.linkText("Funds"))).click().perform();	
							WebElement Sublink15 = w.findElement(By.linkText("Equity/Derivatives"));
							Sublink15.click();
										
							//capture utilisation total
							w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).click();
							
							///check scrip and capture amount
							
							if(w.getPageSource().contains(scrip1))
							{
								///get util amount of specific scrip				
								String total = w.findElement(By.xpath("//td[text()='"+scrip1+"']/following-sibling::td[6]")).getText();
								System.out.println("mod order total " +total);
								tot = Double.parseDouble(total);
								System.out.println("mod order cap amt " +tot);
								Double match = tot - cap;
								System.out.println("mod order match " +match);
								matchround = Math.round(match*100.0)/100.0;
								System.out.println("mod order mar util matchround" +matchround);
								//log.log(LogStatus.INFO, "Margin Utilised Cash & FNO " +matchround);
								
								
								rw = sw.getRow(i);
								cell = rw.createCell(variables.wcutil);
								rw.removeCell(cell);
								cell.setCellValue(tot);
								
								//rw = sw.getRow(i);
								//cell = rw.createCell(11);
								//rw.removeCell(cell);
								//cell.setCellValue(matchround);
																
							////to check if reqd margin and captured util amt are same
							if(matchround.equals(roundoff1))
							{
							///click on cancel
							w.findElement(By.partialLinkText("CANCEL")).click();
							
							///click on total utilised cashNfno
							String totalmargin = w.findElement(By.xpath("//*[@id=\'limit-node-1\']/td[3]")).getText();
							System.out.println(totalmargin);
							String t = totalmargin.replaceAll(",","").trim();
							String m = t.replaceAll("","");
							System.out.println("mod order totalmargin " +m);
							log.log(LogStatus.INFO, "Total Margin is " +totalmargin);
							
							rw = sw.getRow(i);
							cell = rw.createCell(variables.wmtotmar);
							rw.removeCell(cell);
							cell.setCellValue(totalmargin);
							
							///click on margin utilised cashNfno
							String marginutilised = w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).getText();
							System.out.println(marginutilised);
							String t1 = marginutilised.replaceAll(",","").trim();
							String m1 = t1.replaceAll("","");
							System.out.println("mod order mar utilised" +m1);
							log.log(LogStatus.INFO, "Margin Utilised is " +marginutilised);
							
							rw = sw.getRow(i);
							cell = rw.createCell(variables.wmtotmarutil);
							rw.removeCell(cell);
							cell.setCellValue(marginutilised);
							
							///click on margin available cashNfno
							String marginavailable = w.findElement(By.xpath("//*[@id=\'limit-node-6\']/td[3]/strong")).getText();
							System.out.println(marginavailable);
							String t2 = marginavailable.replaceAll(",","").trim();
							String m2 = t2.replaceAll("","");
							System.out.println("mod order mar available " +m2);
							log.log(LogStatus.INFO, "Margin Available is " +marginavailable);
							
							rw = sw.getRow(i);
							cell = rw.createCell(variables.wmmaravail);
							rw.removeCell(cell);
							cell.setCellValue(marginavailable);
							
							//verify totalmargin - margin utilsed = margin available
							Double ab = Double.parseDouble(m)-Double.parseDouble(m1);
							BigDecimal bd = new BigDecimal(ab);
							System.out.println("bd " +bd);
							DecimalFormat df = new DecimalFormat("0.00");
							df.setMaximumFractionDigits(2);
							sd = df.format(ab);
							System.out.println("mod order mar util compare " +sd);
							
							rw = sw.getRow(i);
							cell = rw.createCell(variables.wmactmaravail);
							rw.removeCell(cell);
							cell.setCellValue(sd);
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
								String capturemsg1 = w.findElement(By.xpath("//*[@id=\'container\']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
								log.log(LogStatus.ERROR, "" +capturemsg1);
								w.findElement(By.partialLinkText("BACK")).click();
								//w.findElement(By.xpath(".//input[starts-with(@value,'"+stckno+"')]")).click();
							}
						  }
						}
						else
						{
							System.out.println("Order no not generated as unable to place order.");
							log.log(LogStatus.ERROR, "Order no not generated as " +popup1);
						}
				  	
					  	}	
	  				
	  				
					  	///cancel order
						log = report1.startTest("Cancel " +forreportscripname+ " Order");
						  
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
							if(stckno !=null &&  !stckno.isEmpty())
							{
							if(w.getPageSource().contains(stckno))
							{
							w.findElement(By.xpath(".//input[starts-with(@value,'"+stckno+"')]")).click();
							log.log(LogStatus.INFO, "Click on Order No " +stckno);
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
								System.out.println("Order no "+stckno+ " cannot be cancelled as status is " +status);
								
								newamt = cap;
								System.out.println("newamt " +newamt);
								
								rw = sw.getRow(i);
								cell = rw.createCell(variables.wclastutil);
								cell.setCellValue(newamt);
							}
							else
							{
								//cnfrm
								w.findElement(By.partialLinkText("Confirm")).click();
								log.log(LogStatus.INFO, "Click on Confirm");
								System.out.println("Order no "+stckno+ " is cancelled successfully");
								log.log(LogStatus.INFO, "Order no " +stckno+ " is cancelled successfully");
										
								newamt = cap;
								rw = sw.getRow(i);
								cell = rw.createCell(variables.wclastutil);
								cell.setCellValue(newamt);
								
								///call capture total
								j=5;
								System.out.println("j new "+j);
								capturetotal();
								Double canamt = newamt - cap;
								System.out.println("can amt " +canamt);
								log.log(LogStatus.INFO, "Cancelled Required Margin is " +canamt);
										
								//Actions act15 = new Actions(w);
								//act15.moveToElement(w.findElement(By.linkText("Funds"))).click().perform();	
								//WebElement Sublink15 = w.findElement(By.linkText("Equity/Derivatives"));
								//Sublink15.click();
											
								//capture utilisation total
								//w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).click();
								
								if(w.getPageSource().contains(scrip1))
								{
									///get util amount of specific scrip				
									String total = w.findElement(By.xpath("//td[text()='"+scrip1+"']/following-sibling::td[6]")).getText();
									System.out.println("can order total " +total);
									tot = Double.parseDouble(total);
									System.out.println("can order cap amt " +tot);
									Double match = tot - cap;
									System.out.println("can order match " +match);
									matchround = Math.round(match*100.0)/100.0;
									System.out.println("can order matchround" +matchround);
																										
								////to check if reqd margin and captured util amt are same
								if(matchround.equals(canamt))
								{
								///click on cancel
								w.findElement(By.partialLinkText("CANCEL")).click();
								
								///click on total utilised cashNfno
								String totalmargin = w.findElement(By.xpath("//*[@id=\'limit-node-1\']/td[3]")).getText();
								System.out.println(totalmargin);
								String t = totalmargin.replaceAll(",","").trim();
								String m = t.replaceAll("","");
								System.out.println("total can order " +m);
								log.log(LogStatus.INFO, "Total Margin is " +totalmargin);
								
								rw = sw.getRow(i);
								cell = rw.createCell(variables.wctotmar);
								rw.removeCell(cell);
								cell.setCellValue(totalmargin);
								
								///click on margin utilised cashNfno
								String marginutilised = w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).getText();
								System.out.println(marginutilised);
								String t1 = marginutilised.replaceAll(",","").trim();
								String m1 = t1.replaceAll("","");
								System.out.println("can order mar util " +m1);
								log.log(LogStatus.INFO, "Margin Utilised is " +marginutilised);
								
								rw = sw.getRow(i);
								cell = rw.createCell(variables.wctotmarutil);
								rw.removeCell(cell);
								cell.setCellValue(marginutilised);
								
								///click on margin available cashNfno
								String marginavailable = w.findElement(By.xpath("//*[@id=\'limit-node-6\']/td[3]/strong")).getText();
								System.out.println(marginavailable);
								String t2 = marginavailable.replaceAll(",","").trim();
								String m2 = t2.replaceAll("","");
								System.out.println("can order mar avail " +m2);
								log.log(LogStatus.INFO, "Margin Available is " +marginavailable);
								
								rw = sw.getRow(i);
								cell = rw.createCell(variables.wcmaravail);
								rw.removeCell(cell);
								cell.setCellValue(marginavailable);
								
								//verify totalmargin - margin utilsed = margin available
								Double ab = Double.parseDouble(m)-Double.parseDouble(m1);
								BigDecimal bd = new BigDecimal(ab);
								System.out.println("bd " +bd);
								DecimalFormat df = new DecimalFormat("0.00");
								df.setMaximumFractionDigits(2);
								sd = df.format(ab);
								System.out.println("can order mar util capmare " +sd);
								
								rw = sw.getRow(i);
								cell = rw.createCell(variables.wcactmaravail);
								rw.removeCell(cell);
								cell.setCellValue(sd);
								}
								}
							}
						  
						  }
							else
							{
								System.out.println("Order no not generated as unable to place order.");
								log.log(LogStatus.ERROR, "Order no not generated as " +popup1);
							}
						  }
						k=j;
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
						
		
	@Test(priority=45)
	public void nouse()
	  {
		  report1.endTest(log);
		  report1.flush();
		  w.get("D:\\Saleema\\looputil.html");
	  }
}
