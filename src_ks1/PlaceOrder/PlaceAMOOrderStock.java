package PlaceOrder;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.LogStatus;

import Main.BSOrderXL;

public class PlaceAMOOrderStock extends BSOrderXL
{
	public String stckno;
	public String Status1;
	public String status;
	//public String popup;
	public static ExtentReports report1;
	
	@Test(priority = 12)
	  public void PlaceAMOBuyOrder() throws Exception 
	  {	
			report1 = new ExtentReports("D:\\Saleema\\AMOStockreport.html");
			log = report1.startTest("Place AMO Buy Order for Stock");
			
		  	w.findElement(By.id("userid")).sendKeys("1Y040");
		  	log.log(LogStatus.INFO, "Enter Username");
		  	w.findElement(By.id("pwd")).sendKeys("login1");
		  	log.log(LogStatus.INFO, "Enter Password");
			w.findElement(By.id("otp")).sendKeys("1111");
			log.log(LogStatus.INFO, "Enter Security Key");
			w.findElement(By.partialLinkText("Login")).click();
			log.log(LogStatus.INFO, "Logged in Successfully");
			System.out.println("Logged in Successfully");
			
			///click place order
			Actions act = new Actions(w);
			act.moveToElement(w.findElement(By.linkText("Place Order"))).click().perform();	
			WebElement Sublink = w.findElement(By.linkText("Stock F&O"));
			Sublink.click();
			log.log(LogStatus.INFO, "Click on Place Order >> Stock F&O");

			////select buy option
			w.findElement(By.xpath(".//input[@value='BUY']")).click();
			log.log(LogStatus.INFO, "Select Value");
					
			///select market exchange
			Select ME = new Select(w.findElement(By.id("market_exchange")));
			ME.selectByVisibleText("NSE-Derivative");
			log.log(LogStatus.INFO, "Select Market Exchange");
			
			///select stock name
			w.findElement(By.id("ssscrip")).sendKeys("acc");
			log.log(LogStatus.INFO, "Enter Script");
			w.findElement(By.partialLinkText("ACC LIMITED")).click();
			log.log(LogStatus.INFO, "Select Script");
					
			///select expiry date
			Select ED = new Select(w.findElement(By.id("ssdrvexp")));
			ED.selectByVisibleText("26APR18");
			log.log(LogStatus.INFO, "Select Expiry Date");
			
			///select instrument type
			Select IT = new Select(w.findElement(By.id("ssdrvit")));
			IT.selectByVisibleText("FS");
			log.log(LogStatus.INFO, "Select Instrument Type");
			
			//enter no of lots
			w.findElement(By.id("stk_lot")).sendKeys("2");
			log.log(LogStatus.INFO, "Enter No of Lots : 2");
			
			//price
			//w.findElement(By.id("stk_price")).clear();
					
			//w.findElement(By.id("stk_price")).sendKeys("250");
					
			///check AMO
			w.findElement(By.xpath("//*[@id=\"amo_div\"]/input")).click();;
			log.log(LogStatus.INFO, "Select AMO Checkbox");
			
			///place order
			w.findElement(By.partialLinkText("PLACE ORDER")).click();
			log.log(LogStatus.INFO, "Click on Place Order button");
			
			///alert
			w.switchTo().alert().accept();	
			log.log(LogStatus.INFO, "Click OK on alert");
			
			String popup = "ACTION NOT ALLOWED IN CURRENT PHASE";
			if(w.getPageSource().contains(popup))
			{
				String popup1 = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
				System.err.println("Error displayed as " +popup1);
				log.log(LogStatus.INFO, "Error displayed as " +popup1);
			}
			else
			{
				String stckno = w.findElement(By.xpath("//*[@id='container']/div/p/span")).getText();
				log.log(LogStatus.INFO, "Order No is " +stckno);
				System.out.println(stckno);
				
			}
	  }
	
	  @Test(priority = 13)
	  public void ManageAMOBuyOrder() throws Exception
	  {
		  log = report1.startTest("Manage AMO Buy Order for Stock");
		  String amocheck = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
		  	if(w.getPageSource().contains(amocheck))
		  	{
		  		log.log(LogStatus.INFO, "Error as " +amocheck);
		  	}
		  	else
		  	{
		  		System.out.println("Manage AMO Buy Code is pending");
		  	}
		  
	  }
	  
	  @Test(priority = 14)
	  public void CancelAMOBuyOrder() throws Exception
	  {
		  log = report1.startTest("Cancel AMO Buy Order for Stock");
		  String amocheck = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
		  if(w.getPageSource().contains(amocheck))
		  	{
		  		log.log(LogStatus.INFO, "Error as " +amocheck);
		  	}
		  else
		  {
			  System.out.println("Cancel AMO Buy Code is pending");
		  }
	  }
	
	@Test(priority = 15)
	public void PlaceAMOSellOrder() throws Exception 
	  {		
		log = report1.startTest("Place AMO Sell Order for Stock");
		
			///click place order
			Actions act = new Actions(w);
			act.moveToElement(w.findElement(By.linkText("Place Order"))).click().perform();	
			WebElement Sublink = w.findElement(By.linkText("Stock F&O"));
			Sublink.click();
			
			////select buy option
			w.findElement(By.xpath(".//input[@value='SELL']")).click();
					
			///select market exchange
			Select ME = new Select(w.findElement(By.id("market_exchange")));
			ME.selectByVisibleText("NSE-Derivative");
										
			///select stock name
			w.findElement(By.id("ssscrip")).sendKeys("dlf");
			w.findElement(By.partialLinkText("DLF LIMITED")).click();
					
			///select expiry date
			Select ED = new Select(w.findElement(By.id("ssdrvexp")));
			ED.selectByVisibleText("26APR18");
			
			///select instrument type
			Select IT = new Select(w.findElement(By.id("ssdrvit")));
			IT.selectByVisibleText("FS");
			
			//enter no of lots
			w.findElement(By.id("stk_lot")).sendKeys("2");
							
			//price
			//w.findElement(By.id("stk_price")).clear();
					
			//w.findElement(By.id("stk_price")).sendKeys("250");
					
			///check AMO
			w.findElement(By.xpath("//*[@id=\"amo_div\"]/input")).click();;
								
			///place order
			w.findElement(By.partialLinkText("PLACE ORDER")).click();
					
			///alert
			w.switchTo().alert().accept();	
					
			String popup = "ACTION NOT ALLOWED IN CURRENT PHASE";
			if(w.getPageSource().contains(popup))
			{
				String popup1 = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
				System.err.println("Error displayed as " +popup1);
			}
			else
			{
				String stckno = w.findElement(By.xpath("//*[@id='container']/div/p/span")).getText();
				System.out.println(stckno);
				
			}
	  }
	
	  @Test(priority = 16)
	  public void ManageAMOSellOrder() throws Exception
	  {
		  log = report1.startTest("Manage AMO Sell Order for Stock");
		  System.out.println("Manage AMO Sell Code is pending");
	  }
	  
	  @Test(priority = 17)
	  public void CancelAMOSellOrder() throws Exception
	  {
		  log = report1.startTest("Cancel AMO Sell Order for Stock");
		  System.out.println("Cancel AMO Sell Code is pending");
	  }
	  
	  @Test(priority = 3000)
	  public void nouse2()
	  {
		  report1.endTest(log);
		  report1.flush();
		  w.get("D:\\Saleema\\AMOStockreport.html");
	  }
}
