package Utilisation;

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

public class DerivativeUtilisation extends BSOrderXL
{
	public static String capture;
	public static Double cap;
	public static Double matchround;
	public static Double roundoff;
	public static Double newroundoff;
	public static Double price;
	public static Double round;
	public static Double tot;
	public static Double RM;
	public static String stc;
	public static int multiple;
	public static int nooflots;
	public static int quantity;
	public static String sd;
	public static String scrip1;
	public String stckno;
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
					
		//capture utilisation total
		w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).click();
		String checkmsg = "No Record Found";
		String scrip1 = s.getRow(7).getCell(3).getStringCellValue();
		if(!w.getPageSource().contains(checkmsg))
		{
			if(w.getPageSource().contains(scrip1))
			{
				String capture = w.findElement(By.xpath("//td[text()='"+scrip1+"']/following-sibling::td[6]")).getText();
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
	
	@Test(priority=0)
	public void PlaceAMOBuyOrder() throws Exception 
	  {
		FileInputStream f = new FileInputStream("D:\\Saleema\\KotakSecurities\\src\\uticash.xls");
		HSSFWorkbook wb = new HSSFWorkbook(f);
		HSSFSheet s = wb.getSheet("Utilisation");
		HSSFSheet s1 = wb.getSheet("Watchlist");
		
		///report creation
		report1 = new ExtentReports("D:\\Saleema\\DerivativeUtilisation.html");
		log = report1.startTest("Manage Watchlist");
	  	
		////call login function
	  	Login();
	  	
	  	///call manage watchlist
	  	//cm.manage();
	  	
	  	log = report1.startTest("Place Cash Order Derivative");
	  	
	  	///get multiple value from excel for utilization caln
	  	multiple = (int) s.getRow(3).getCell(3).getNumericCellValue();  	
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
		String xcelscrip = s.getRow(5).getCell(3).getStringCellValue();
		System.out.println("data1 " +xcelscrip);
		w.findElement(By.xpath(".//input[@value='"+xcelscrip+"']")).click();
		
		////click on place order
		w.findElement(By.partialLinkText("PLACE ORDER")).click();
		log.log(LogStatus.INFO, "Click on Place Order");
		
		///get scripname
		String scripname = w.findElement(By.id("stk_name")).getAttribute("value");
		log.log(LogStatus.INFO, "Selected Scrip is " +scripname);
		
		///select buy/sell
		String value = s.getRow(6).getCell(3).getStringCellValue();
		w.findElement(By.xpath(".//input[@value='"+value+"']")).click();
		log.log(LogStatus.INFO, "Selected Value is " +value);
		
		////enter quantity
		int nooflots = (int) s.getRow(4).getCell(3).getNumericCellValue();  	
		System.out.println("data1 " +nooflots);
		//wb.close();
		w.findElement(By.id("stk_lot")).sendKeys(String.valueOf(nooflots));
		log.log(LogStatus.INFO, "Entered No of lots is: " +nooflots);
				
		///required margin calculation
		String pr1 = w.findElement(By.id("stk_price")).getAttribute("value");
		w.findElement(By.id("stk_price")).clear();
		price = (Double.parseDouble(pr1)*1)/100;
		double pr = Double.parseDouble(pr1);
		double ans = pr-price;
		round = (double) (Math.round(ans));
		w.findElement(By.id("stk_price")).sendKeys(String.valueOf(round));
		log.log(LogStatus.INFO, "Actual price is " +pr1);
		log.log(LogStatus.INFO, "Modified price is " +round);
		String qty = w.findElement(By.id("stk_qty")).getAttribute("value");
		log.log(LogStatus.INFO, "Quantity is: " +qty);
		quantity = Integer.parseInt(qty);
		RM = ((quantity * round ) / multiple);
		roundoff = (Math.round(RM*100.0)/100.0);
		log.log(LogStatus.INFO, "Required Margin is " +roundoff);
		System.out.println("roundoff " +roundoff);
		
		///chk for AMO
		String amochk = s.getRow(8).getCell(3).getStringCellValue();
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
			//w.findElement(By.partialLinkText("BACK")).click();
		}
		else
		{	
			////capture order no
			stckno = w.findElement(By.xpath("//*[@id='container']/div/p/span")).getText();
			log.log(LogStatus.INFO, "Order No is " +stckno);
			
			///click funds >>equity/derivatives
			Actions act1 = new Actions(w);
			act1.moveToElement(w.findElement(By.linkText("Funds"))).click().perform();	
			WebElement Sublink1 = w.findElement(By.linkText("Equity/Derivatives"));
			Sublink1.click();
			log.log(LogStatus.INFO, "Click on Funds >> Equity/Derivatives");
			
			//click margin utilised cash&FNO
			w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).click();
			
			///check scrip and capture amount
			String scrip1 = s.getRow(7).getCell(3).getStringCellValue();
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
				else
				{
					log.log(LogStatus.ERROR, "Total Margin & Margin utilized are not equal to Margin Available " +sd);
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
				System.out.println("Selected scrip does not match");
			}
	  }
						
	}
	
	@Test(priority=1)
	public void ManageCashAMO() throws Exception 
	  {		
		///report creation
		log = report1.startTest("Change Cash Order Derivative");
	  	
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
		Actions act = new Actions(w);
		act.moveToElement(w.findElement(By.linkText("Reports"))).click().perform();	
		WebElement Sublink = w.findElement(By.linkText("Order Status"));
		Sublink.click();
		log.log(LogStatus.INFO, "Click on Reports >> Order Status");
		
	  	//	w.findElement(By.partialLinkText("Order Status")).click();	
	  	
		//check stkno
		if(stckno !=null &&  !stckno.isEmpty())
		{
		w.findElement(By.xpath(".//input[@value='"+stckno+"|Y']")).click();
		log.log(LogStatus.INFO, "Click on Order No " +stckno);
		
		//check status of order no
		String status = w.findElement(By.xpath("//td[text()='"+stckno+"']/following-sibling::td[6]")).getText();
		if(status.equals("AMO") || status.equals("OPN"))
		{
		///click change order
		w.findElement(By.partialLinkText("CHANGE ORDER")).click();
				
		String check = "Error";
		if(w.getPageSource().contains(check))
		{
			String popup = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
			log.log(LogStatus.ERROR, "" +popup);
			//w.findElement(By.partialLinkText("BACK")).click();
			System.out.println("Order no " +stckno+ " cannot be modified as status is " +status);

		}
		else
		{		
			FileInputStream f = new FileInputStream("D:\\Saleema\\KotakSecurities\\src\\uticash.xls");
			HSSFWorkbook wb = new HSSFWorkbook(f);
			HSSFSheet s = wb.getSheet("Utilisation");
			
			Double newcap = matchround;
									
			//update stock no
			w.findElement(By.id("stk_lot")).clear();
			nooflots = (int) s.getRow(4).getCell(8).getNumericCellValue();  	
			w.findElement(By.id("stk_lot")).sendKeys(String.valueOf(nooflots));
			log.log(LogStatus.INFO, "Modified No of lots : " +nooflots);
			//wb.close();
			
			w.findElement(By.id("stk_price")).click();
			
			String qt1 = w.findElement(By.id("stk_qty")).getAttribute("value");
			quantity = Integer.parseInt(qt1);
			log.log(LogStatus.INFO, "Quantity is : " +quantity);
			///modify order
			RM = ((quantity * round ) / multiple);
			Double roundoff1 = (Math.round(RM*100.0)/100.0);
			System.out.println("roundoff " +roundoff1);
			
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
			String chkerror = "Error";
			if(w.getPageSource().contains(chkerror))
			{
				String capturemsg1 = w.findElement(By.xpath("//*[@id=\'container\']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
				log.log(LogStatus.ERROR, "" +capturemsg1);
				w.findElement(By.partialLinkText("BACK")).click();
				w.findElement(By.xpath(".//input[@value='"+stckno+"|N']")).click();
			}
		}
		}
		else
		{
			System.out.println("Order no not generated as unable to place order.");
			log.log(LogStatus.INFO, "Order no not generated as unable to place order.");
		}
	  	}
  }
		
	 @Test(priority = 2)
	  public void CancelAMOBuyOrder() throws Exception
	  {
		 ///for report name
		  log = report1.startTest("Cancel Cash Order Derivative");
		  
		  ///check for any error
		  String amocheck = "Error";
		  if(w.getPageSource().contains(amocheck))
		  	{
			  	String capturemsg = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
		  		log.log(LogStatus.ERROR, "" +capturemsg);
		   	}
		  else
		  {
			///click watchlist>>place order
			Actions act = new Actions(w);
			act.moveToElement(w.findElement(By.linkText("Reports"))).click().perform();	
			WebElement Sublink = w.findElement(By.linkText("Order Status"));
			Sublink.click();
			log.log(LogStatus.INFO, "Click on Reports >> Order Status");
			
			//check stkno
			if(stckno !=null &&  !stckno.isEmpty())
			{
				w.findElement(By.xpath(".//input[@value='"+stckno+"|Y']")).click();
				log.log(LogStatus.INFO, "Click on Order no " +stckno);
				
				///click on cancel
				w.findElement(By.partialLinkText("CANCEL ORDER")).click();
				log.log(LogStatus.INFO, "Click on Cancel");
				
				///check for error
				String trade = "Can not cancel Order.";
				if(w.getPageSource().contains(trade))
				{
					String popup = w.findElement(By.xpath("//*[@id=\'container\']/div/form/table[1]/tbody/tr[2]/td[7]")).getText();
					log.log(LogStatus.ERROR, "Order no " +stckno+ " cannot be cancelled");
					log.log(LogStatus.ERROR, "" +popup);				
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
									
					///call capture total
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
	  }
	 
	@Test(priority=47)
	public void nouse()
	  {
		  report1.endTest(log);
		  report1.flush();
		  w.get("D:\\Saleema\\DerivativeUtilisation.html");
	  }
}
