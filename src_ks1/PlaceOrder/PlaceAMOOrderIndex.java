package PlaceOrder;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import Main.BSOrderXL;

public class PlaceAMOOrderIndex extends BSOrderXL
{
	public static ExtentReports report1;
	public static ExtentTest log;
	public String stckno;
	public String Status1;
	public String status;

	@Test(priority = 0)
	  public void PlaceAMOBuyOrder() throws Exception 
	  {	
			report1 = new ExtentReports("D:\\Saleema\\AMOIndexreport.html");
			log = report1.startTest("Place AMO Buy Order for Index");
						
		  	w.findElement(By.id("userid")).sendKeys("1Y040");
		  	log.log(LogStatus.INFO, "Enter Username");
			w.findElement(By.id("pwd")).sendKeys("login1");
			log.log(LogStatus.INFO, "Enter password");
			w.findElement(By.id("otp")).sendKeys("1111");
			log.log(LogStatus.INFO, "Enter Access Code");
			w.findElement(By.partialLinkText("Login")).click();
			log.log(LogStatus.INFO, "Logged in successfully");
			System.out.println("Logged in Successfully");
			
			///click place order
			Actions act = new Actions(w);
			act.moveToElement(w.findElement(By.linkText("Place Order"))).click().perform();	
			WebElement Sublink = w.findElement(By.linkText("Index F&O"));
			Sublink.click();
			log.log(LogStatus.INFO, "Click on Place order >> Index F&O");
			////select buy option
			w.findElement(By.xpath(".//input[@value='BUY']")).click();
			log.log(LogStatus.INFO, "Select value");
					
			///select market exchange
			Select ME = new Select(w.findElement(By.id("market_exchange")));
			ME.selectByVisibleText("NSE-Derivative");
			log.log(LogStatus.INFO, "Select Market Exchange : NSE-Derivative");
			
			///select stock name
			w.findElement(By.id("ssscrip")).sendKeys("nif");
			w.findElement(By.partialLinkText("NIFTY")).click();
			log.log(LogStatus.INFO, "Enter Stock Name : NIFTY");
			
			///select expiry date
			Select ED = new Select(w.findElement(By.id("ssdrvexp")));
			ED.selectByVisibleText("26APR18");
			log.log(LogStatus.INFO, "Select Expiry Date");
			
			///select instrument type
			Select IT = new Select(w.findElement(By.id("ssdrvit")));
			IT.selectByVisibleText("FI");
			log.log(LogStatus.INFO, "Select Instrument type");
			
			//enter no of lots
			w.findElement(By.id("stk_lot")).sendKeys("4");
			log.log(LogStatus.INFO, "Enter No of Lots : 4");
			
			//price
			//w.findElement(By.id("stk_price")).clear();
					
			//w.findElement(By.id("stk_price")).sendKeys("250");
					
			///check AMO
			w.findElement(By.xpath("//*[@id=\"amo_div\"]/input")).click();
			log.log(LogStatus.INFO, "Select AMO checkbox");
								
			///place order
			w.findElement(By.partialLinkText("PLACE ORDER")).click();
			log.log(LogStatus.INFO, "Click Place Order button");
			
			///alert
			w.switchTo().alert().accept();	
			log.log(LogStatus.INFO, "Click OK on Alert message");
			
			String popup = "ACTION NOT ALLOWED IN CURRENT PHASE";
			if (w.getPageSource().contains(popup))
			{
				String popup1 = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
				System.err.println("Error displayed as " +popup1);
				log.log(LogStatus.INFO, "Error displayed as " +popup1);
			}
			else
			{
				String stckno = w.findElement(By.xpath("//*[@id='container']/div/p/span")).getText();
				System.out.println(stckno);
				log.log(LogStatus.INFO, "Order No is " +stckno);
				
			}
			
	  }
	
	  @Test(priority = 1)
	  public void ManageAMOBuyOrder() throws Exception
	  {
		  	log = report1.startTest("Manage AMO Buy Order for Index");
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
	  
	  @Test(priority = 2)
	  public void CancelAMOBuyOrder() throws Exception
	  {
		  log = report1.startTest("Cancel AMO Buy Order for Index");
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
	
	@Test(priority = 3)
	public void PlaceAMOSellOrder() throws Exception 
	  {
			log = report1.startTest("Place AMO Sell Order for Index");
			///click place order
			Actions act = new Actions(w);
			act.moveToElement(w.findElement(By.linkText("Place Order"))).click().perform();	
			WebElement Sublink = w.findElement(By.linkText("Index F&O"));
			Sublink.click();
			log.log(LogStatus.INFO, "Click on Place Order >> Index F&O");
			
			////select buy option
			w.findElement(By.xpath(".//input[@value='SELL']")).click();
			log.log(LogStatus.INFO, "Select Value");
			
			///select market exchange
			Select ME = new Select(w.findElement(By.id("market_exchange")));
			ME.selectByVisibleText("NSE-Derivative");
			log.log(LogStatus.INFO, "Select Market Exchange");
										
			///select stock name
			w.findElement(By.id("ssscrip")).sendKeys("f");
			w.findElement(By.partialLinkText("FTSE100")).click();
			log.log(LogStatus.INFO, "Enter FTSE100");
					
			///select expiry date
			Select ED = new Select(w.findElement(By.id("ssdrvexp")));
			ED.selectByVisibleText("20APR18");
			log.log(LogStatus.INFO, "Select Expiry Date");
			
			///select instrument type
			Select IT = new Select(w.findElement(By.id("ssdrvit")));
			IT.selectByVisibleText("FI");
			log.log(LogStatus.INFO, "Select Instrument Type");
			
			//enter no of lots
			w.findElement(By.id("stk_lot")).sendKeys("3");
			log.log(LogStatus.INFO, "Enter No of Lots : 3");
							
			//price
			//w.findElement(By.id("stk_price")).clear();
					
			//w.findElement(By.id("stk_price")).sendKeys("250");
					
			///check AMO
			w.findElement(By.xpath("//*[@id=\"amo_div\"]/input")).click();
			log.log(LogStatus.INFO, "Select AMO Checkbox");
								
			///place order
			w.findElement(By.partialLinkText("PLACE ORDER")).click();
			log.log(LogStatus.INFO, "Click on Place Order button");
			
			///alert
			w.switchTo().alert().accept();	
			log.log(LogStatus.INFO, "Click OK on Alert");
			
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
				System.out.println(stckno);
				log.log(LogStatus.INFO, "Stock no " +stckno);
			}
	  }
	
	  @Test(priority = 4)
	  public void ManageAMOSellOrder() throws Exception
	  {
		  log = report1.startTest("Manage AMO Sell Order for Index");
		  System.out.println("Manage AMO Sell Code is pending");
	  }
	  
	  @Test(priority = 5)
	  public void CancelAMOSellOrder() throws Exception
	  {
		  log = report1.startTest("Cancel AMO Sell Order for Index");
		  System.out.println("Cancel AMO Sell Code is pending");
	  }
	  
	  @Test(priority = 2000)
	  public void nouse1()
	  {
		  report1.endTest(log);
		  report1.flush();
		  w.get("D:\\Saleema\\AMOIndexreport.html");
	  }
	  
}
