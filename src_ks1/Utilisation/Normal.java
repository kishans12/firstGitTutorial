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

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.LogStatus;

import Main.BSOrderXL;

public class Normal extends BSOrderXL
{
	public static String capture;
	public static String scripname;
	public static Double cap;
	public static Double matchround;
	public static Double roundoff;
	public static Double roundoff1;
	public static Double newroundoff;
	public static Double price;
	public static Double roundoffpr;
	public static Double tot;
	public static Double RM;
	public static String stc;
	public static int multiple;
	public static int quantity;
	public static String sd;
	public static String scrip1;
	public String stckno;
	public String Status1;
	public String status;
	public static int j;
	public static int i;
	public static int k;
	public static Row rw;
	public static Cell cell;
	
	Utilisation.ManageWatchlist cm = new Utilisation.ManageWatchlist();
	
	public void capturetotal() throws Exception
	{
		//FileInputStream f = new FileInputStream(srcfile);
		//wb = new HSSFWorkbook(f);
		HSSFSheet s = wb.getSheet("Sheet1");
		
		///click funds >>equity/derivatives
		Actions act1 = new Actions(w);
		act1.moveToElement(w.findElement(By.linkText("Funds"))).click().perform();	
		WebElement Sublink1 = w.findElement(By.linkText("Equity/Derivatives"));
		Sublink1.click();
					
		//capture utilisation total
		w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).click();
		String checkmsg = "No Record Found";
		
		System.out.println("j cap "+j);

		scrip1 = s.getRow(i).getCell(j).getStringCellValue();
		System.out.println("scrip1 "+scrip1);
		if(!w.getPageSource().contains(checkmsg))
		{
			if(w.getPageSource().contains(scrip1))
			{
			System.out.println("in loop scrip1 " +scrip1);
			String capture = w.findElement(By.xpath("//td[text()='"+scrip1+"']/following-sibling::td[6]")).getText();
			System.out.println("capture " +capture);
			cap = Double.parseDouble(capture);
			System.out.println("first/cap " +cap);
			}
			else
			{
				cap = 0.0;
				System.out.println("cap of capture " +cap);
			}
			rw = s.getRow(i);
			cell = rw.createCell(9);
			rw.removeCell(cell);
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
		HSSFSheet s = wb.getSheet("Sheet1");
		HSSFSheet s1 = wb.getSheet("Sheet2");

		///report creation
		report1 = new ExtentReports("D:\\Saleema\\looputil.html");
		log = report1.startTest("Manage Watchlist");
	  	
		////call login function
	  	Login();
	  	
	  	///call manage watchlist
	  	//cm.manage();
	  		  	
		//log = report1.startTest("Place Normal " +watchlistname+ " Cash Order");
	  	
	  	///get multiple value from excel for utilization caln

	  	int rc = s.getLastRowNum() - s.getFirstRowNum();
	  	System.out.println("rc "+rc);
	  	String watchlistname = s1.getRow(1).getCell(1).getStringCellValue();
	  	
	  	for(i=1;i<rc+1;i++)
	  	{
	  		Row row = s.getRow(i);
	  		System.out.println("k "+k);
	  		System.out.println("j "+j);
	  				for(j=0;j<row.getLastCellNum();)
	  				{
	  					System.out.println("start i"+i);
	  					System.out.println("start j"+j);
	  					
					  	String forreportscripname=s.getRow(i).getCell(j).getStringCellValue();
					  	System.out.println("reportscripname "+forreportscripname);
					  	log = report1.startTest("Place Normal " +forreportscripname+ " Cash Order");
					  	
					  	j=j+1;
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

						////click on place order
						w.findElement(By.partialLinkText("PLACE ORDER")).click();
						log.log(LogStatus.INFO, "Click on Place Order");
						
						///get scripname
						scripname = w.findElement(By.id("stk_name")).getAttribute("value");
						log.log(LogStatus.INFO, "Selected Scrip is " +scripname);
						
						///select buy/sell
						j=j+1;
						String value = s.getRow(i).getCell(j).getStringCellValue();
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
							multiple = (int) s.getRow(i).getCell(j).getNumericCellValue();  	
							System.out.println("data " +multiple);
							
							///required margin calculation
							j=j+1;
							int perc = (int) s.getRow(i).getCell(j).getNumericCellValue();
							String pr1 = w.findElement(By.id("stk_price")).getAttribute("value");
							w.findElement(By.id("stk_price")).clear();
							price = (Double.parseDouble(pr1)*perc)/100;
							double chngestrngpr = Double.parseDouble(pr1);
							double total = chngestrngpr-price;
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
							multiple = (int) s.getRow(i).getCell(j).getNumericCellValue();  	
							System.out.println("data " +multiple);
							
							///required margin calculation
							j=j+1;
							int perc = (int) s.getRow(i).getCell(j).getNumericCellValue();
							String pr1 = w.findElement(By.id("stk_price")).getAttribute("value");
							w.findElement(By.id("stk_price")).clear();
							price = (Double.parseDouble(pr1)*perc)/100;
							double chngestrngpr = Double.parseDouble(pr1);
							double total = chngestrngpr-price;
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
							String popup1 = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
							System.err.println("Error displayed as " +popup1);
							log.log(LogStatus.ERROR, "" +popup1);
							
						}
						else
						{	
							Thread.sleep(1000);
							////capture order no
							stckno = w.findElement(By.xpath("//*[@id='container']/div/p/span")).getText();
							System.out.println("stckno...... " +stckno);
							log.log(LogStatus.INFO, "Order No is " +stckno);
							
							System.out.println("i......"+i);
							rw = s.getRow(i);
							cell = rw.createCell(10);
							rw.removeCell(cell);
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
								
								rw = s.getRow(i);
								cell = rw.createCell(11);
								rw.removeCell(cell);
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
								
								rw = s.getRow(i);
								cell = rw.createCell(12);
								rw.removeCell(cell);
								cell.setCellValue(totalmargin);
								
								///click on margin utilised cashNfno
								String marginutilised = w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).getText();
								System.out.println(marginutilised);
								String t1 = marginutilised.replaceAll(",","").trim();
								String m1 = t1.replaceAll("","");
								System.out.println(m1);
								log.log(LogStatus.INFO, "Margin Utilised is " +marginutilised);
								
								rw = s.getRow(i);
								cell = rw.createCell(13);
								rw.removeCell(cell);
								cell.setCellValue(marginutilised);
								
								///click on margin available cashNfno
								String marginavailable = w.findElement(By.xpath("//*[@id=\'limit-node-6\']/td[3]/strong")).getText();
								System.out.println(marginavailable);
								String t2 = marginavailable.replaceAll(",","").trim();
								String m2 = t2.replaceAll("","");
								System.out.println(m2);
								log.log(LogStatus.INFO, "Margin Available is " +marginavailable);
								
								rw = s.getRow(i);
								cell = rw.createCell(14);
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
								
								rw = s.getRow(i);
								cell = rw.createCell(15);
								rw.removeCell(cell);
								cell.setCellValue(sd);
								
								///check if above are equal
								if(sd.equals(m2))
								{
									System.out.println("Equal");
									log.log(LogStatus.INFO, "Utilized -- As Required Margin and Margin Utilised are Equal");
									log.log(LogStatus.INFO, "Total Margin & Margin utilized are equal to Margin Available " +sd);
									
									rw = s.getRow(i);
									cell = rw.createCell(16);
									rw.removeCell(cell);
									cell.setCellValue("Pass");
								}
								else
								{
									log.log(LogStatus.ERROR, "Total Margin & Margin utilized are not equal to Margin Available "+sd);
									
									rw = s.getRow(i);
									cell = rw.createCell(16);
									rw.removeCell(cell);
									cell.setCellValue("Fail");
								}
								
							}
							else
							{
								log.log(LogStatus.ERROR, "Not Utilized -- As Required Margin and Margin Utilised are Unequal");
								log.log(LogStatus.ERROR, "Total Margin & Margin utilized are not equal to Margin Available ");
								
								rw = s.getRow(i);
								cell = rw.createCell(16);
								rw.removeCell(cell);
								cell.setCellValue("Fail");
							}
						}
							else
							{
								System.out.println("Selected scrip does not match");
							}
							
							FileOutputStream Of = new FileOutputStream(srcfile);
							wb.write(Of);
							Of.close();
					  }
						
						///manage order
						log = report1.startTest("Change Normal " +forreportscripname+ " Cash Order");
						
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
							//FileInputStream f = new FileInputStream("D:\\Saleema\\KotakSecurities\\newxl.xls");
							//HSSFWorkbook wb = new HSSFWorkbook(f);
							//HSSFSheet s = wb.getSheet("Sheet1");
							
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
							log.log(LogStatus.INFO, "Enter Quantity: " +quantity);
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
								//wb.close();
								
								w.findElement(By.id("stk_price")).click();
								
								String qt1 = w.findElement(By.id("stk_qty")).getAttribute("value");
								quantity = Integer.parseInt(qt1);
								log.log(LogStatus.INFO, "Quantity is : " +quantity);
								///modify order
								RM = ((quantity * roundoffpr ) / multiple);
								roundoff1 = (Math.round(RM*100.0)/100.0);
								System.out.println("roundoff " +roundoff1);
							}
							///click on chnge order
							w.findElement(By.partialLinkText("CHANGE ORDER")).click();
							log.log(LogStatus.INFO, "Click on Change Order");	
							
							//cnfrm
							w.findElement(By.partialLinkText("Confirm")).click();
							log.log(LogStatus.INFO, "Click Confirm");
							System.out.println("Order no " +stckno+ " is modified successfully");
							log.log(LogStatus.INFO, "Order no " +stckno+ " is modified successfully");
														
							log.log(LogStatus.INFO, "Modified required margin is " +roundoff1);
													
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
								w.findElement(By.xpath(".//input[starts-with(@value,'"+stckno+"')]")).click();
							}
						  }
						}
						else
						{
							System.out.println("Order no not generated as unable to place order.");
							log.log(LogStatus.INFO, "Order no not generated as unable to place order.");
						}
				  	
					  	}	
	  				
	  				
					  	///cancel order
						log = report1.startTest("Cancel Normal " +forreportscripname+ " Cash Order");
						  
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
							}
							else
							{
								//cnfrm
								w.findElement(By.partialLinkText("Confirm")).click();
								log.log(LogStatus.INFO, "Click on Confirm");
								System.out.println("Order no "+stckno+ " is cancelled successfully");
								log.log(LogStatus.INFO, "Order no " +stckno+ " is cancelled successfully");
								
								Double newamt = cap;
								System.out.println("newamt " +newamt);
								
								///call capture total
								j=1;
								System.out.println("j new "+j);
								capturetotal();
								Double canamt = newamt - cap;
								System.out.println("can amt " +canamt);
								log.log(LogStatus.INFO, "Cancelled Required Margin is " +canamt);
							}
						  
						  }
							else
							{
								System.out.println("Order no not generated as unable to place order.");
								log.log(LogStatus.INFO, "Order no not generated as unable to place order.");
							}
						  }
						k=j;
						System.out.println("end i"+i);
				  		System.out.println("end j"+j);
				  		System.out.println("end k"+k);
				  		break;
					}
	  		}
	  	}
		
	@Test(priority=45)
	public void nouse()
	  {
		  report1.endTest(log);
		  report1.flush();
		  w.get("D:\\Saleema\\looputil.html");
	  }
}
