package TSLO;

import java.io.FileInputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.LogStatus;

import Main.BSOrderXL;

public class BSEUtilisation extends BSOrderXL
{
	public static String capture;
	public static String capturemsg;
	public static Double cap;
	public static String msg;
	public static Double matchround;
	public static Double roundoff;
	public static Double newroundoff;
	public static Double price;
	public static Double tot;
	public static Double RM;
	public static String stc;
	public static int multiple;
	public static int quantity;
	public static String sd;
	public static String scrip1;
	public static int orderno;
	public String Status1;
	public String status;
	
	Utilisation.ManageWatchlist cm = new Utilisation.ManageWatchlist();
			
	public void capturetotal() throws Exception
	{
		FileInputStream f = new FileInputStream("D:\\Saleema\\KotakSecurities\\src\\uticash.xls");
		HSSFWorkbook wb = new HSSFWorkbook(f);
		HSSFSheet s = wb.getSheet("Utilisation");
		
		///click funds >>equity/derivatives
		Actions act1 = new Actions(w);
		act1.moveToElement(w.findElement(By.linkText("Funds"))).click().perform();	
		WebElement Sublink1 = w.findElement(By.linkText("Equity/Derivatives"));
		Sublink1.click();
		
		String message = "Error";
		if(w.getPageSource().contains(message))
		{
			String popup = w.findElement(By.xpath("//*[@id=\'container\']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
			System.out.println("cap message "+popup);
			log.log(LogStatus.INFO, "" +popup);
			//w.findElement(By.partialLinkText("BACK")).click();
		}
		else
		{
		//capture utilisation total
		w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).click();
		String checkmsg = "No Record Found";
		String scrip1 = s.getRow(7).getCell(2).getStringCellValue();
		if(!w.getPageSource().contains(checkmsg))
		{
			if(w.getPageSource().contains(scrip1))
			{
				String capture = w.findElement(By.xpath("//td[text()='"+scrip1+"']/following-sibling::td[6]")).getText();
				System.out.println("capture " +capture);
				cap = Double.parseDouble(capture);
				System.out.println("first/cap " +cap);
			}
			else
			{
				cap = 0.0;
				System.out.println("first/cap " +cap);
			}
		}
		else
		{
			System.out.println("Out of loop");
		}
		}		
	}
	
	@Test(priority=0)
	public void PlaceAMOBuyOrder() throws Exception 
	  {
		FileInputStream f = new FileInputStream("D:\\Saleema\\KotakSecurities\\src\\uticash.xls");
		HSSFWorkbook wb = new HSSFWorkbook(f);
		HSSFSheet s = wb.getSheet("Utilisation");
		HSSFSheet s1 = wb.getSheet("Watchlist");
		
		///report creation
		//report1 = new ExtentReports("D:\\Saleema\\TSLOUtilisation.html");
		//log = report1.startTest("Manage Watchlist");
	  	
		////call login function
	  	//Login();
	  	
	  	///call manage watchlist
	  	//cm.manage();
	  	
	  	log = report1.startTest("Place TSLO Order BSE");
	  	
	  	///get multiple value from excel for utilization caln
	  	multiple = (int) s.getRow(3).getCell(2).getNumericCellValue();  	
		System.out.println("data " +multiple);
				
		///call capture total function to capture utilisation value
		capturetotal();
		
	  	///click watchlist>>place order
		Actions act = new Actions(w);
		act.moveToElement(w.findElement(By.linkText("Watchlist"))).click().perform();	
		WebElement Sublink = w.findElement(By.linkText("My Watchlists"));
		Sublink.click();
		log.log(LogStatus.INFO, "Click on Watchlist >> My Watchlists");
		
		///watchlistname
		String watchlistname = s1.getRow(1).getCell(1).getStringCellValue();
		w.findElement(By.xpath("//div[@id='container']//a[contains(text(), '"+watchlistname+"')]")).click();
				
		///select scrip
		String xcelscrip = s.getRow(5).getCell(2).getStringCellValue();
		System.out.println("data1 " +xcelscrip);
		w.findElement(By.xpath(".//input[@value='"+xcelscrip+"']")).click();
		
		////click on place order
		w.findElement(By.partialLinkText("PLACE ORDER")).click();
		log.log(LogStatus.INFO, "Click on Place Order");
		
		///get scripname
		String scripname = w.findElement(By.id("stk_name")).getAttribute("value");
		log.log(LogStatus.INFO, "Selected Scrip is " +scripname);
		
		///select buy/sell
		String value = s.getRow(6).getCell(2).getStringCellValue();
		w.findElement(By.xpath(".//input[@value='"+value+"']")).click();
		log.log(LogStatus.INFO, "Selected Value is " +value);
		
		////enter quantity
		int quantity = (int) s.getRow(4).getCell(2).getNumericCellValue();  	
		System.out.println("data1 " +quantity);
		//wb.close();
		w.findElement(By.id("stk_qty")).sendKeys(String.valueOf(quantity));
		log.log(LogStatus.INFO, "Entered Quantity is: " +quantity);
				
		///required margin calculation
		String pr1 = w.findElement(By.id("stk_price")).getAttribute("value");
		w.findElement(By.id("stk_price")).clear();
		price = Double.parseDouble(pr1)-10;
		w.findElement(By.id("stk_price")).sendKeys(String.valueOf(price));
		log.log(LogStatus.INFO, "Actual price is " +pr1);
		log.log(LogStatus.INFO, "Modified price is " +price);
		RM = ((quantity * price ) / multiple);
		roundoff = (Math.round(RM*100.0)/100.0);
		log.log(LogStatus.INFO, "Required Margin is " +roundoff);
		System.out.println("roundoff " +roundoff);
				
		///chk for TSLO
		String tslochk = s.getRow(10).getCell(2).getStringCellValue();
		System.out.println("TSLOchk "+tslochk);
		if(tslochk.equals("YES"))
		{
			//chk checkbox
			w.findElement(By.xpath(".//input[@value= '6']")).click();
			log.log(LogStatus.INFO, "Select TSLO New checkbox");
					
			///enter spread value
			int spread = (int) s.getRow(11).getCell(2).getNumericCellValue();
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
		
		///check for any error
		String popup = "Error";
		if (w.getPageSource().contains(popup))
		{
			String popup1 = w.findElement(By.xpath("//*[@id=\'container\']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
			System.err.println("Error displayed as " +popup1);
			log.log(LogStatus.ERROR, "" +popup1);
			//w.findElement(By.partialLinkText("BACK")).click();
		}
		else
		{	
			////capture order no
			msg = w.findElement(By.xpath("//*[@id='container']/div[1]/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
			System.out.println("msg "+msg);
			log.log(LogStatus.INFO, "" +msg);
					
			///get order no
			orderno = Integer.parseInt(msg.replaceAll("\\D", ""));
			System.out.println("order no "+orderno);
					
			w.findElement(By.partialLinkText("SHOW TSLO STATUS")).click();
			
			w.findElement(By.xpath(".//input[@value='"+orderno+",3']")).click();
						
			String status = w.findElement(By.xpath("//td[text()='"+orderno+"']/following-sibling::td[10]")).getText();
			Thread.sleep(3000);
			
			System.out.println("status "+status);
			if(!status.equals("OPN"))
			{
					log.log(LogStatus.ERROR, "Status is " +status+ " so margin utilised cannot be calculated");
			}
			else
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
				String scrip1 = s.getRow(7).getCell(2).getStringCellValue();
				System.out.println("scrip1 "+scrip1);
				if(w.getPageSource().contains(scrip1))
				{
					///get util amount of specific scrip				
					String total = w.findElement(By.xpath("//td[text()='"+scrip1+"']/following-sibling::td[6]")).getText();
					tot = Double.parseDouble(total);
					System.out.println("cap amt " +tot);
					Double match = tot - cap;
					matchround = Math.round(match*100.0)/100.0;
					System.out.println("f amt" +matchround);
					log.log(LogStatus.INFO, "Margin calculated is " +matchround);
				
				////to check if reqd margin and captured util amt are same
				if(matchround.equals(roundoff))
				{
					///click on cancel
					w.findElement(By.partialLinkText("CANCEL")).click();
					
					///click on total utilised cashNfno
					String totalmargin = w.findElement(By.xpath("//*[@id=\'limit-node-1\']/td[3]")).getText();
					String t = totalmargin.replaceAll(",","").trim();
					String m = t.replaceAll("","");
					System.out.println(m);
					log.log(LogStatus.INFO, "Total Margin is " +totalmargin);
				
					///click on margin utilised cashNfno
					String marginutilised = w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).getText();
					String t1 = marginutilised.replaceAll(",","").trim();
					String m1 = t1.replaceAll("","");
					System.out.println(m1);
					log.log(LogStatus.INFO, "Margin Utilised is " +marginutilised);
					
					///click on margin available cashNfno
					String marginavailable = w.findElement(By.xpath("//*[@id=\'limit-node-6\']/td[3]/strong")).getText();
					String t2 = marginavailable.replaceAll(",","").trim();
					String m2 = t2.replaceAll("","");
					System.out.println(m2);
					log.log(LogStatus.INFO, "Margin Available is " +marginavailable);
					
					//verify totalmargin - margin utilsed = margin available
					Double ab = Double.parseDouble(m)-Double.parseDouble(m1);
					BigDecimal bd = new BigDecimal(ab);
					DecimalFormat df = new DecimalFormat("0.00");
					df.setMaximumFractionDigits(2);
					sd = df.format(ab);
					System.out.println(sd);
					
					///check if above are equal
					if(sd.equals(m2))
					{
						System.out.println("Equal");
						log.log(LogStatus.INFO, "Utilized -- As Required Margin and Margin Utilised are Equal");
						log.log(LogStatus.INFO, "Total Margin & Margin utilized are equal to Margin Available " +sd);
					}
				
				}
				else
				{
					log.log(LogStatus.ERROR, "Not Utilized -- As Required Margin and Margin Utilised are Unequal");
					log.log(LogStatus.ERROR, "Total Margin & Margin utilized are not equal to Margin Available ");
				}
			}
				else
				{
					System.out.println("Selected scrip is not available");
				}
		  }
			
		}
						
	}
	
	@Test(priority=1)
	public void ManageCashAMO() throws Exception 
	  {		
		///report creation
		log = report1.startTest("Change TSLO Order BSE");
	  	
		///to check for any error
		String amocheck = "Error";
	  	if(w.getPageSource().contains(amocheck))
	  	{
	  		String capturemsg = w.findElement(By.xpath("//*[@id=\'container\']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
	  		log.log(LogStatus.ERROR, "" +capturemsg);
	  	}
	  	else
	  	{
		  	///click watchlist>>place order
			Actions act = new Actions(w);
			act.moveToElement(w.findElement(By.linkText("Reports"))).click().perform();	
			WebElement Sublink = w.findElement(By.linkText("TSLO Status"));
			Sublink.click();
			log.log(LogStatus.INFO, "Click on Reports >> TSLO Status");
			
			//if(msg !=null &&  !msg.isEmpty())
			//{
		  	w.findElement(By.xpath(".//input[@value='"+orderno+",3']")).click();
			log.log(LogStatus.INFO, "Click on TSLO Order No " +orderno);
			
			//check status of order no
			String status = w.findElement(By.xpath("//td[text()='"+orderno+"']/following-sibling::td[10]")).getText();
			if(status.equals("ACTIVE") || status.equals("OPN"))
			{
				///click change order
				w.findElement(By.partialLinkText("MODIFY ORDER")).click();
				Thread.sleep(2000);
				
				String check = "BS Indicator";
				if(!w.getPageSource().contains(check))
				{
					System.out.println("main loop");
					String popup = w.findElement(By.xpath("//*[@id='container']/div[1]/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
					log.log(LogStatus.ERROR, "" +popup);
					//w.findElement(By.partialLinkText("BACK")).click();
					System.out.println("Order no " +orderno+ " cannot be modified as status is " +status);
				}
				else
				{		
					FileInputStream f = new FileInputStream("D:\\Saleema\\KotakSecurities\\src\\uticash.xls");
					HSSFWorkbook wb = new HSSFWorkbook(f);
					HSSFSheet s = wb.getSheet("Utilisation");
					
					Double newcap = matchround;
											
					//update stock no
					w.findElement(By.id("leg1_ord_qty")).clear();
					quantity = (int) s.getRow(4).getCell(7).getNumericCellValue();  	
					w.findElement(By.id("leg1_ord_qty")).sendKeys(String.valueOf(quantity));
					log.log(LogStatus.INFO, "Modified Quantity : " +quantity);
					//wb.close();
					
					///modify order
					RM = ((quantity * price ) / multiple);
					Double roundoff1 = (Math.round(RM*100.0)/100.0);
					System.out.println("roundoff " +roundoff1);
					
					///click on chnge order
					w.findElement(By.partialLinkText("MODIFY")).click();
					log.log(LogStatus.INFO, "Click on Modify ");	
					
					//cnfrm
					w.findElement(By.partialLinkText("CONFIRM")).click();
					log.log(LogStatus.INFO, "Click Confirm");
					
					String msg1 = w.findElement(By.xpath("//*[@id='container']/div[1]/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
					log.log(LogStatus.INFO, "" +msg1);
					
					System.out.println("Order no " +orderno+ " is modified successfully");
					log.log(LogStatus.INFO, "Order no " +orderno+ " is modified successfully");
					
					log.log(LogStatus.INFO, "Modified required margin is " +roundoff1);
				}
			}
			else
			{
				w.findElement(By.partialLinkText("MODIFY ORDER")).click();
				Thread.sleep(2000);
				
				String chkerror = "Message";
				if(w.getPageSource().contains(chkerror))
				{
					String capturemsg1 = w.findElement(By.xpath("//*[@id='container']/div[1]/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
					log.log(LogStatus.ERROR, "" +capturemsg1);
					w.findElement(By.partialLinkText("BACK")).click();
					w.findElement(By.xpath(".//input[@value='"+orderno+",3']")).click();
				}
			}
			//}
		//else
		//{
			//System.out.println("Order no not generated as unable to place order.");
			//log.log(LogStatus.INFO, "Order no not generated as unable to place order.");
		//}
	  	}
  }
		
	 @Test(priority = 2)
	  public void CancelAMOBuyOrder() throws Exception
	  {
		 ///for report name
		  log = report1.startTest("Cancel TSLO Order BSE");
		  
		  ///check for any error
		  String amocheck = "Error";
		  if(w.getPageSource().contains(amocheck))
		  {
				String capturemsg = w.findElement(By.xpath("//*[@id=\'container\']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
			  	log.log(LogStatus.ERROR, "" +capturemsg);
		  }
		  else
		  {
				///click watchlist>>place order
				Actions act = new Actions(w);
				act.moveToElement(w.findElement(By.linkText("Reports"))).click().perform();	
				WebElement Sublink = w.findElement(By.linkText("TSLO Status"));
				Sublink.click();
				log.log(LogStatus.INFO, "Click on Reports >> TSLO Status");
			
				//if(msg !=null &&  !msg.isEmpty())
				//{	
				w.findElement(By.xpath(".//input[@value='"+orderno+",3']")).click();
				log.log(LogStatus.INFO, "Click on Order no " +orderno);
				
				///click on cancel
				w.findElement(By.partialLinkText("CANCEL ORDER")).click();
				log.log(LogStatus.INFO, "Click on Cancel");
				Thread.sleep(2000);
				
				///check for error
				String trade = "BS Indicator";
				if(!w.getPageSource().contains(trade))
				{
					String popup = w.findElement(By.xpath("//*[@id='container']/div[1]/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
					log.log(LogStatus.ERROR, "" +popup);				
					w.findElement(By.partialLinkText("BACK")).click();	
					System.out.println("Order no "+orderno+ " cannot be cancelled as status is " +status);
				}
				else
				{
					//cnfrm
					w.findElement(By.partialLinkText("CANCEL ORDER")).click();
					log.log(LogStatus.INFO, "Click on Cancel Order ");
					System.out.println("Order no "+orderno+ " is cancelled successfully");
					log.log(LogStatus.INFO, "Order no " +orderno+ " is cancelled successfully");
					
					Double newamt = cap;
										
					///call capture total
					capturetotal();
					Double canamt = newamt - cap;
					System.out.println("can amt " +canamt);
					log.log(LogStatus.INFO, "Cancelled Required Margin is " +canamt);
				}
			}
	  
			//else
			//{
			//	System.out.println("Order no not generated as unable to place order.");
			//	log.log(LogStatus.INFO, "Order no not generated as unable to place order.");
			//}
		  //}
	  }
	 
	@Test(priority=52)
	public void nouse()
	  {
		// report1.endTest(log);
		 //report1.flush();
		 //w.get("D:\\Saleema\\TSLOUtilisation.html");
	  }
	  
}
