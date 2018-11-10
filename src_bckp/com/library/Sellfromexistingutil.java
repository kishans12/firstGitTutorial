package com.library;

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
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

public class Sellfromexistingutil
{
	  public static WebDriver w;
	  public static HSSFWorkbook wb;
	  public static HSSFSheet s;
	  public static HSSFSheet s2;
	  public static ExtentReports report1;
	  public static ExtentTest log;
	  String URL1;
	  String Uname;
	  String Pwd;
	  int AccCode;
	  int i,j,rc;
	  int column;
	  public static String symbol,stkupld;
	  public static String status;
	  public static String celtext;
	  public static WebElement ele;
	  	  
	  com.library.variables vr = new com.library.variables();
	  com.library.functions fn = new com.library.functions();
	  
	@BeforeSuite
	  public void beforeSuite1() throws Exception
	  {	  				
		  	System.setProperty("webdriver.chrome.driver", "D:\\Selenium software\\chromedriver.exe");
			w = new ChromeDriver();
			
			FileInputStream f = new FileInputStream(vr.srcfile);
			wb = new HSSFWorkbook(f);
			s = wb.getSheet("Config");
			
			/// get URL
			URL1 = s.getRow(2).getCell(1).getStringCellValue();
			w.get(URL1);
			w.manage().window().maximize();
			w.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
			
	  }
	
	 public void Login() throws IOException
	  {	  		
			///enter username
		  	Uname = s.getRow(4).getCell(2).getStringCellValue();
		  	w.findElement(By.id("userid")).sendKeys(Uname);
		  	log.log(LogStatus.INFO, "Enter User name " +Uname);
		  	
		  	///enter password
		  	Pwd  = s.getRow(4).getCell(3).getStringCellValue();
			w.findElement(By.id("pwd")).sendKeys(Pwd);
			log.log(LogStatus.INFO, "Enter Password ******");
			
			///enter access code
			AccCode= (int) s.getRow(4).getCell(4).getNumericCellValue();
			w.findElement(By.id("otp")).sendKeys(String.valueOf(AccCode));
			log.log(LogStatus.INFO, "Enter Security Key/Access Code ****");
			
			///click on login
			w.findElement(By.partialLinkText("Login")).click();
			System.out.println("Logged in Successfully");
			log.log(LogStatus.INFO, "Logged in Successfully");
			
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
					
			//symbol = s2.getRow(i).getCell(j).getStringCellValue();
			String space= " ";			
			vr.combine = symbol.toUpperCase().concat(space).concat("EQ").concat(space).concat(vr.forreportscripname);
			vr.scrip1 = vr.combine;
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
		public void PlaceSFEOrder() throws Exception 
		  {
			 	FileInputStream f = new FileInputStream(vr.srcfile);
				wb = new HSSFWorkbook(f);
			  				  	
				s2 = wb.getSheet("Stockupload");
				
				///report creation
				report1 = new ExtentReports("D:\\Saleema\\sfeorder.html");
				log = report1.startTest("Login");
								
				////call login function
				Login();
				
				rc = s2.getLastRowNum() - s2.getFirstRowNum();
				System.out.println("rc "+rc);
				
				for(i=5;i<rc+1;i++)
			  	{
			  		Row row = s2.getRow(i);
			  		 
			  		String flag = s2.getRow(i).getCell(0).getStringCellValue();
						
						if(flag.equals("Y"))
						{		
							symbol = s2.getRow(i).getCell(2).getStringCellValue();
							System.out.println("symbol " +symbol);
							
			  				for(j=4;j<row.getLastCellNum();)
			  				{	
			  					//stckbal = (int) s2.getRow(i).getCell(j).getNumericCellValue();			  						  					
			  					
			  					vr.forreportscripname=s2.getRow(i).getCell(j).getStringCellValue();
							  	System.out.println("reportscripname "+vr.forreportscripname);
							  	log = report1.startTest("Place " +vr.forreportscripname+ " Sell from Existing Order");							  	  			  					  			  		
							  												  								  					
							  	///click watchlist>>place order
								Actions act = new Actions(w);
								act.moveToElement(w.findElement(By.linkText("Place Order"))).click().perform();	
								WebElement Sublink = w.findElement(By.linkText("Sell From Existing"));
								Sublink.click();
								log.log(LogStatus.INFO, "Click on Place Order >> Sell From Existing");
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////								
								j=j+1;
			  					stkupld=s2.getRow(i).getCell(j).getStringCellValue();
			  					if(stkupld.equals("DEL"))
			  					{
									w.findElement(By.partialLinkText("Del")).click();
									log.log(LogStatus.INFO, "Click on Del Mark");
			  					}
			  					else if(stkupld.equals("BNSTG"))
			  					{
			  						w.findElement(By.partialLinkText("BNST")).click();
									log.log(LogStatus.INFO, "Click on BNSTG Mark");
			  					}
			  					else
			  					{
			  						w.findElement(By.partialLinkText("HOLD")).click();
									log.log(LogStatus.INFO, "Click on HOLD Mark");
			  					}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			  					if(w.getPageSource().contains(symbol.toUpperCase()))
			  					{		  					
			  					
								String getscrip =w.findElement(By.xpath("//div[@id='container']//td[contains(text(), '"+symbol.toUpperCase()+"')]")).getText();
								System.out.println("getsc "+getscrip);																			
																	
									WebElement mytable = w.findElement(By.xpath("//div[@id='container']/div/form/div/table/tbody"));
							    	//To locate rows of table. 
							    	List < WebElement > rows_table = mytable.findElements(By.tagName("tr"));
							    	//To calculate no of rows In table.
							    	int rows_count = rows_table.size();
							    	//Loop will execute till the last row of table.
							    	for (int row1 = 4; row1 < rows_count;) 
							    	{
							    	    //To locate columns(cells) of that specific row.
							    	    List < WebElement > Columns_row = rows_table.get(row1).findElements(By.tagName("td"));
							    	    //To calculate no of columns (cells). In that specific row.
							    	    int columns_count = Columns_row.size();			
							    	    //System.out.println("Number of cells In Row " + row1 + " are " + columns_count);
							    	    
							    	    //Loop will execute till the last cell of that specific row.
							    	    for (column = 0; column < columns_count;) 
							    	    {
							    	        // To retrieve text from that specific cell.
							    	        celtext = Columns_row.get(column).getText();
							    	        //System.out.println("Cell Value of row number " + row1 + " and column number " + column + " Is " + celtext);
							    	        break;
							    	    }

							    	        if(celtext.contains(getscrip) && vr.forreportscripname.equals("NSE"))
							    	        {						    	        	
							    	        	int col = column +3;
							    	        	w.findElement(By.xpath("//*[@id='container']/div/form/div/table/tbody/tr[" + (row1+1)+ "]/td[" + (col)+ "]/input")).click();
							    	        	break;
							    	        }
							    	        else if(celtext.contains(getscrip) && vr.forreportscripname.equals("BSE"))
							    	        {
							    	        	int col = column + 1;
							    	        	w.findElement(By.xpath("//*[@id='container']/div/form/div/table/tbody/tr[" + (row1+2)+ "]/td[" + (col)+ "]/input")).click();
							    	        	break;
							    	        }	
							    	        
							    	        row1=row1+3;
							    	    }						    	    
							    	
								//w.findElement(By.partialLinkText("Place Order")).click();
								w.findElement(By.xpath("//*[@id='container']/div/form/a[2]")).click();
								
								j=j+1;
								vr.quantity = (int) s2.getRow(i).getCell(j).getNumericCellValue();
								w.findElement(By.id("delqty00")).sendKeys(String.valueOf(vr.quantity));
								log.log(LogStatus.INFO, "Entered Quantity is: " +vr.quantity);													
								
								///multiple
								j=j+1;
								vr.multiple = (float) s2.getRow(i).getCell(j).getNumericCellValue();  	
								System.out.println("multiple " +vr.multiple);
								
								///get bal qty
								String stckbal = w.findElement(By.xpath("//*[@id='container']/div/form/div/table/tbody/tr[3]/td[6]")).getText();
			  					Double cnvrtstring = Double.parseDouble(stckbal);
			  					
								///required margin calculation
								j=j+1;
								int perc = (int) s2.getRow(i).getCell(j).getNumericCellValue();
								vr.pr1 = w.findElement(By.id("delprice00")).getAttribute("value");
								w.findElement(By.id("delprice00")).clear();
								vr.price = (Double.parseDouble(vr.pr1)*perc)/100;
								double chngestrngpr = Double.parseDouble(vr.pr1);
								vr.total = chngestrngpr+vr.price;
								vr.roundoffpr = (double) (Math.round(vr.total));
								w.findElement(By.id("delprice00")).sendKeys(String.valueOf(vr.roundoffpr));
								log.log(LogStatus.INFO, "Actual price is " +vr.pr1);
								log.log(LogStatus.INFO, "Modified price is " +vr.roundoffpr);
								vr.RM = ((cnvrtstring * vr.roundoffpr ) / vr.multiple);
								//vr.roundoff = (Math.round(vr.RM*100.0)/100.0);
								vr.roundoff = vr.RM;
								System.out.println("roundoff " +vr.roundoff);
								
								///write actual price value to excel
								j=j+1;
								vr.rw = s2.getRow(i);
								vr.cell = vr.rw.createCell(j);
								vr.cell.setCellValue(vr.pr1);
								
								///write modified price value to excel
								j=j+1;
								vr.rw = s2.getRow(i);
								vr.cell = vr.rw.createCell(j);
								vr.cell.setCellValue(vr.roundoffpr);
								
								//w.findElement(By.partialLinkText("Place Order")).click();
								w.findElement(By.xpath("//*[@id='container']/div/form/table[2]/tbody/tr/td/a[1]")).click();
/////if qty exceeds								
								w.findElement(By.partialLinkText("Confirm")).click();
																								
								Thread.sleep(1000);
								////capture order no
								vr.stckno = w.findElement(By.xpath("//*[@id='container']/div/table[2]/tbody/tr[3]/td[3]")).getText();
								System.out.println("stckno...... " +vr.stckno);
								log.log(LogStatus.INFO, "Order No is " +vr.stckno);
								
								if(vr.stckno.equals("-NA-"))
								{
									vr.stckno=null;
									vr.popup1 = w.findElement(By.xpath("//*[@id='container']/div/table[2]/tbody/tr[3]/td[4]")).getText();
								}
								capturetotal();
								
								//to check place util is not 0
								if(vr.cap!=0)
								{
									j=j+1;
									vr.rw = s2.getRow(i);
									vr.cell = vr.rw.createCell(j);
									vr.cell.setCellValue(vr.cap);
							  		log.log(LogStatus.ERROR, "Utilisation is " +vr.cap);

									break;
								}
								///write placed utilisation value to excel
								j=j+1;
								vr.rw = s2.getRow(i);
								vr.cell = vr.rw.createCell(j);
								vr.cell.setCellValue(vr.cap);
						  		log.log(LogStatus.INFO, "Utilisation is " +vr.cap);
						  		
								///manage order
								log = report1.startTest("Change " +vr.forreportscripname+ " Order");
								
								///to check for any error
								String amocheck = "Error";
							  	if(w.getPageSource().contains(amocheck))
							  	{
							  		String capturemsg = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
							  		log.log(LogStatus.ERROR, "" +capturemsg);
							  		
							  		vr.rw = s2.getRow(i);
									vr.cell = vr.rw.createCell(variables.sfestatus);
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
								status = w.findElement(By.xpath("//td[text()='"+vr.stckno+"']/following-sibling::td[6]")).getText();
								if(status.equals("AMO") || status.equals("OPN"))
								{
												
								///click change order
								w.findElement(By.partialLinkText("CHANGE ORDER")).click();
										
								String check = "Error";
								if(w.getPageSource().contains(check))
								{
									vr.popup1 = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
									log.log(LogStatus.ERROR, "" +vr.popup1);
									//w.findElement(By.partialLinkText("BACK")).click();
									System.out.println("Order no " +vr.stckno+ " cannot be modified as status is " +status);
									
									vr.rw = s2.getRow(i);
									vr.cell = vr.rw.createCell(variables.sfestatus);
									vr.cell.setCellValue(vr.popup1);
								}
								else
								{														
									Double newcap = vr.cap;
									System.out.println("newcap " +newcap);							
																						
									//update stock no
									if(w.findElement(By.id("stk_qty")).isEnabled())
									{
									j=j+1;
									w.findElement(By.id("stk_qty")).clear();
									vr.quantity = (int) s2.getRow(i).getCell(j).getNumericCellValue();  	
									System.out.println("data1 " +vr.quantity);
									w.findElement(By.id("stk_qty")).sendKeys(String.valueOf(vr.quantity));
									log.log(LogStatus.INFO, "Modified Quantity is : " +vr.quantity);																									
									
									///modify order
									vr.RM = ((vr.quantity * vr.roundoffpr ) / vr.multiple);
									System.out.println("RM " +vr.RM);
									//vr.roundoff1 = (Math.round(vr.RM*100.0)/100.0);
									vr.roundoff1 = vr.RM;
									System.out.println("roundoff " +vr.roundoff1);
									}
									else
									{
										j=j+1;
										w.findElement(By.id("stk_lot")).clear();
										vr.quantity = (int) s2.getRow(i).getCell(j).getNumericCellValue();  	
										w.findElement(By.id("stk_lot")).sendKeys(String.valueOf(vr.quantity));
										log.log(LogStatus.INFO, "Modified No of lots : " +vr.quantity);
																		
										w.findElement(By.id("stk_price")).click();
										
										String qt1 = w.findElement(By.id("stk_qty")).getAttribute("value");
										vr.quantity = Integer.parseInt(qt1);
										log.log(LogStatus.INFO, "Quantity is : " +vr.quantity);
										///modify order
										vr.RM = ((vr.quantity * vr.roundoffpr ) / vr.multiple);
										//vr.roundoff1 = (Math.round(vr.RM*100.0)/100.0);
										vr.roundoff1 = vr.RM;
										System.out.println("roundoff " +vr.roundoff1);
									}
									
									double util = vr.roundoff1 - vr.roundoff;
									double ro = (Math.round(util*100.0)/100.0);
									System.out.println("util "+ro);
									
									///click on chnge order
									w.findElement(By.partialLinkText("CHANGE ORDER")).click();
									log.log(LogStatus.INFO, "Click on Change Order");	
									
									//cnfrm
									w.findElement(By.partialLinkText("Confirm")).click();
									log.log(LogStatus.INFO, "Click Confirm");
									System.out.println("Order no " +vr.stckno+ " is modified successfully");
									log.log(LogStatus.INFO, "Order no " +vr.stckno+ " is modified successfully");
									log.log(LogStatus.INFO, "Utilisation is " +ro);
																		
									capturetotal();
									Double capamt = vr.cap;
									
									j=j+1;
									vr.rw = s2.getRow(i);
									vr.cell = vr.rw.createCell(j);
									vr.cell.setCellValue(capamt);
									
									if(capamt.equals(ro))
									{
										vr.rw = s2.getRow(i);
										vr.cell = vr.rw.createCell(variables.sfestatus);
										vr.cell.setCellValue("Utilisation and required margin are equal");
									}
									else
									{
										vr.rw = s2.getRow(i);
										vr.cell = vr.rw.createCell(variables.sfestatus);
										vr.cell.setCellValue("Utilisation and required margin are unequal");
									}
									////write util value
									j=j+1;
									vr.rw = s2.getRow(i);
									vr.cell = vr.rw.createCell(j);
									vr.cell.setCellValue(ro);
									
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
										
										vr.rw = s2.getRow(i);
										vr.cell = vr.rw.createCell(variables.sfestatus);
										vr.cell.setCellValue(capturemsg1);
										
										w.findElement(By.partialLinkText("BACK")).click();										
									}
								  }
								}
								else
								{
									System.out.println("Order no not generated as unable to place order.");
									log.log(LogStatus.ERROR, "Order no not generated as " +vr.popup1);
									vr.rw = s2.getRow(i);
									vr.cell = vr.rw.createCell(variables.sfestatus);
									vr.cell.setCellValue(vr.popup1);									
								}						  	
							  }
							  	
							  ///cancel order
								log = report1.startTest("Cancel " +vr.forreportscripname+ " Order");
								  
								  ///check for any error
								  String amocheck1 = "Error";
								  if(w.getPageSource().contains(amocheck1))
								  	{
									  	String capturemsg = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
								  		log.log(LogStatus.ERROR, "" +capturemsg);
								  		
								  		vr.rw = s2.getRow(i);
								  		vr.cell = vr.rw.createCell(variables.sfestatus);
								  		vr.cell.setCellValue(capturemsg);
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
										if(w.getPageSource().contains(trade))
										{
											String popup2 = w.findElement(By.xpath("//*[@id=\'container\']/div/form/table[1]/tbody/tr[2]/td[7]")).getText();
											System.out.println("status is "+popup2);
											log.log(LogStatus.ERROR, "" +popup2);	
											
											vr.rw = s2.getRow(i);
											vr.cell = vr.rw.createCell(variables.sfestatus);
											vr.cell.setCellValue(popup2);
											
											w.findElement(By.partialLinkText("Back")).click();	
											System.out.println("Order no "+vr.stckno+ " cannot be cancelled as status is " +vr.status);																					
										}
										else
										{
											//cnfrm
											w.findElement(By.partialLinkText("Confirm")).click();
											log.log(LogStatus.INFO, "Click on Confirm");
											System.out.println("Order no "+vr.stckno+ " is cancelled successfully");
											log.log(LogStatus.INFO, "Order no " +vr.stckno+ " is cancelled successfully");
										}								  
									}
									else
									{
										System.out.println("Order no not generated as unable to place order.");
										log.log(LogStatus.ERROR, "Order no not generated as " +vr.popup1);
										vr.rw = s2.getRow(i);
										vr.cell = vr.rw.createCell(variables.sfestatus);
										vr.cell.setCellValue(vr.popup1);
									}
								  }
			  					}
			  					else
			  					{
									log.log(LogStatus.INFO, "Stock " +symbol.toUpperCase()+ " didn't got uploaded");
			  					}
								break;
			  				}
						}
						else
						{
							//flag is N
						}
			  	}
				FileOutputStream Of = new FileOutputStream(vr.srcfile);
				wb.write(Of);
				Of.close();
				
				report1.endTest(log);
				report1.flush();
				w.get("D:\\Saleema\\sfeorder.html");
		  }
}
