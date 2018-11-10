package PlaceOrder;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.LogStatus;

import Main.BSOrderXL;

public class PlaceAMOOrderC extends BSOrderXL
{
	public String stckno;
	public String Status1;
	public String status;
	//public String popup;
	public static ExtentReports report1;
	
	public void amocheck() throws Exception
	{
		Date today = Calendar.getInstance().getTime();
		SimpleDateFormat dateFormat = new SimpleDateFormat("hh:mm:ss a");
	    Date time1 = dateFormat.parse("3:00:00 PM");
	    Date time2 = dateFormat.parse("10:00:00 PM");
	    Date time3 = dateFormat.parse("7:00:00 AM");
	    Date time4 = dateFormat.parse("10:00:00 AM");
	    //Date time5 = dateFormat.parse("12:40:50 PM");
	    
	    //Get current time
		Date CurrentTime = Calendar.getInstance().getTime();
	    String date = dateFormat.getTimeInstance().format(CurrentTime);
		System.out.println("date "+date);
		Date ct = dateFormat.parse(date);
		
		if ((ct.after(time1) && ct.before(time2)) || (ct.after(time3) && ct.before(time4))) 
		{
			w.findElement(By.xpath("//*[@id='amo_div']/input")).click();
			log.log(LogStatus.INFO, "Select AMO Checkbox");
		}     
		else 
		{
		    System.out.println("Action not allowed");
		}
		
	}
	
	  @Test(priority = 0)
	  public void PlaceAMOBuyOrder() throws Exception 
	  {
		  	report1 = new ExtentReports("D:\\Saleema\\AMOCashreport.html");
		  	log = report1.startTest("Place AMO Buy Order for Cash");
		  	
		  	w.findElement(By.id("userid")).sendKeys("1Y040");
		  	log.log(LogStatus.INFO, "Enter Username");
			w.findElement(By.id("pwd")).sendKeys("login1");
			log.log(LogStatus.INFO, "Enter password");
			w.findElement(By.id("otp")).sendKeys("1111");
			log.log(LogStatus.INFO, "Enter Security key");
			w.findElement(By.partialLinkText("Login")).click();
			log.log(LogStatus.INFO, "Logged in successfully");
			System.out.println("Logged in Successfully");
			
			///click place order
			Actions act = new Actions(w);
			act.moveToElement(w.findElement(By.linkText("Place Order"))).click().perform();	
			WebElement Sublink = w.findElement(By.linkText("Stocks"));
			Sublink.click();
			log.log(LogStatus.INFO, "Click on Place order >> Stocks");				
			////select buy option
			w.findElement(By.xpath(".//input[@value='BUY']")).click();
			log.log(LogStatus.INFO, "Select Value");
			
			///select market exchange
			Select ME = new Select(w.findElement(By.id("market_exchange")));
			ME.selectByVisibleText("NSE");
			log.log(LogStatus.INFO, "Select Market Exchange");
								
			///select stock name
			w.findElement(By.id("ssscrip")).sendKeys("dlf");
			log.log(LogStatus.INFO, "Enter Script");
			
			//w.findElement(By.cssSelector("DLF LIMITED")).click();
			w.findElement(By.partialLinkText("DLF LIMITED")).click();
			log.log(LogStatus.INFO, "Select Script");
					
			//qty
			w.findElement(By.id("stk_qty")).sendKeys("2");
			log.log(LogStatus.INFO, "Enter Quantity : 2");
			//price
			//w.findElement(By.id("stk_price")).clear();
			
			//w.findElement(By.id("stk_price")).sendKeys("250");
			
			///check AMO
			amocheck();
						
			///place order
			w.findElement(By.partialLinkText("PLACE ORDER")).click();
			log.log(LogStatus.INFO, "Click on Place Order button");
					
			///alert
			w.switchTo().alert().accept();
			log.log(LogStatus.INFO, "Click OK on alert");
			
			String popup = "ACTION NOT ALLOWED IN CURRENT PHASE";
			if (w.getPageSource().contains(popup))
			{
				String popup1 = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
				System.err.println("Error displayed as " +popup1);
				log.log(LogStatus.INFO, "Error displayed as " +popup1);
			}
			else
			{
				stckno = w.findElement(By.xpath("//*[@id='container']/div/p/span")).getText();
				System.out.println(stckno);
				log.log(LogStatus.INFO, "Order No is " +stckno);
														
			}
				  
	  }
	  
	  @Test(priority = 1)
	  public void ManageAMOBuyOrder() throws Exception
	  {
		  	log = report1.startTest("Manage AMO Buy Order for Cash");
		  	String amocheck = "Error";
		  	if(w.getPageSource().contains(amocheck))
		  	{
		  		String capturemsg = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
		  		log.log(LogStatus.INFO, "Error as " +capturemsg);
		  	}
		  	else
		  	{
		  	///order status
			w.findElement(By.partialLinkText("Order Status")).click();
			log.log(LogStatus.INFO, "Click on Order Status");
		  	
			String status = w.findElement(By.xpath("//div[@id='container']//td[contains(text(), 'OPN')]")).getText();
			System.out.println("status is " +status);
			String Status1 = "OPN";				
			System.out.println("Order no" +stckno);
			
			//check stkno
			w.findElement(By.xpath(".//input[@value='"+stckno+"|N']")).click();
			
			///click change order
			w.findElement(By.partialLinkText("CHANGE ORDER")).click();
			
			
			String check = "Input Error";
			if(w.getPageSource().startsWith(check))
			{
			String trade = "You cannot change the Order";
			String popup = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
			if(trade.equals(popup) && status.equals(Status1))
				{
				System.err.println("Error as It is traded " +popup);
				w.findElement(By.partialLinkText("BACK")).click();
				System.out.println("Order no " +stckno+ " cannot be modified as status is " +status);
				}
			}
			else
			{								
				//update stock no
				w.findElement(By.id("stk_qty")).clear();
				
				w.findElement(By.id("stk_qty")).sendKeys("3");
				
				///click on chnge order
				w.findElement(By.partialLinkText("CHANGE ORDER")).click();
								
				//cnfrm
				w.findElement(By.partialLinkText("Confirm")).click();
				System.out.println("Order no " +stckno+ " is modified successfully");
				///order status
				w.findElement(By.partialLinkText("Order Status")).click();
				
				//check stkno
				w.findElement(By.xpath(".//input[@value='"+stckno+"|N']")).click();
			}
		  System.out.println("Manage AMO Buy Code is pending");
		  	}
	  }
	  
	  @Test(priority = 2)
	  public void CancelAMOBuyOrder() throws Exception
	  {
		  log = report1.startTest("Cancel AMO Buy Order for Cash");
		  String amocheck = "Error";
		  if(w.getPageSource().contains(amocheck))
		  	{
			  	String capturemsg = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
		  		log.log(LogStatus.INFO, "Error as " +capturemsg);
		   	}
		  else
		  {
			w.findElement(By.partialLinkText("CANCEL ORDER")).click();
			String trade = "Can not cancel Order.It is already TRADED";
			if(w.getPageSource().contains(trade))
			{
			//String trade = "Can not cancel Order.It is already TRADED";
			String popup = w.findElement(By.xpath("//div[@id='container']//td[contains(text(), 'Can not cancel Order.It is already TRADED')]")).getText();
			System.out.println("status is "+popup);
							
			if(trade.equals(popup) && status.equals(Status1))
				{
				
				System.err.println("Error as It is traded " +popup);
				w.findElement(By.partialLinkText("Back")).click();	
				System.out.println("Order no "+stckno+ " cannot be cancelled as status is " +status);
				
				}
			}
			else
			{
				//cnfrm
				w.findElement(By.partialLinkText("Confirm")).click();
				System.out.println("Order no "+stckno+ " is cancelled successfully");			
			}
		  System.out.println("Cancel AMO Buy Code is pending");
		  }
	  }
	  /*	  
	  @Test(priority = 3)
	  public void PlaceAMOSellOrder() throws Exception
	  {	
		  	log = report1.startTest("Place AMO Sell Order for Cash");
		  
			///click place order
			Actions act = new Actions(w);
			act.moveToElement(w.findElement(By.linkText("Place Order"))).click().perform();	
			WebElement Sublink = w.findElement(By.linkText("Stocks"));
			Sublink.click();
			log.log(LogStatus.INFO, "Click on Place Order >> Stocks");
			
			////select buy option
			w.findElement(By.xpath(".//input[@value='SELL']")).click();
			log.log(LogStatus.INFO, "Select Value");
			
			///select market exchange
			Select ME = new Select(w.findElement(By.id("market_exchange")));
			ME.selectByVisibleText("NSE");
			log.log(LogStatus.INFO, "Select Market Exchange");
					
			///select stock name
			w.findElement(By.id("ssscrip")).sendKeys("dlf");
			log.log(LogStatus.INFO, "Enter Script");
			
			w.findElement(By.partialLinkText("DLF LIMITED")).click();
			log.log(LogStatus.INFO, "Select Script");
			
			//qty
			w.findElement(By.id("stk_qty")).sendKeys("2");
			log.log(LogStatus.INFO, "Enter Quantity : 2");
			
			//price
			//w.findElement(By.id("stk_price")).clear();
			
			//w.findElement(By.id("stk_price")).sendKeys("250");
			
			///check AMO
			w.findElement(By.xpath("//*[@id=\'amo_div\']/input")).click();;
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
				log.log(LogStatus.INFO, "Order No is " +stckno);
											
			}
	  }
	  
	  @Test(priority = 4)
	  public void ManageAMOSellOrder() throws Exception
	  {
		  log = report1.startTest("Change AMO Sell Order for Cash");
		  String amocheck = "ACTION NOT ALLOWED IN CURRENT PHASE";
		  if(w.getPageSource().contains(amocheck))
		  	{
		  		log.log(LogStatus.INFO, "Error as " +amocheck);
		  	}
		  else
		  {
			  System.out.println("Manage AMO Sell Code is pending");
		  }
	  }
	  
	  @Test(priority = 5)
	  public void CancelAMOSellOrder() throws Exception
	  {
		  log = report1.startTest("Cancel AMO Sell Order for Cash");
		  String amocheck = "ACTION NOT ALLOWED IN CURRENT PHASE";
		  if(w.getPageSource().contains(amocheck))
		  	{
		  		log.log(LogStatus.INFO, "Error as " +amocheck);
		  	}
		  else
		  {
			  System.out.println("Cancel AMO Sell Code is pending");
		  }
		  
	  }
	  */
	  @Test(priority = 1000)
	  public void nouse()
	  {
		  report1.endTest(log);
		  report1.flush();
		  w.get("D:\\Saleema\\AMOCashreport.html");
	  }
}
