package PlaceOrder;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.LogStatus;

import Main.BSOrderXL;

public class PlaceNormalOrderC extends BSOrderXL
{
	//public WebDriver w;
	public static String stckno;
	public String Status1;
	public String status;
	
	@Test(priority = 6)
	public void PlaceNormalBuyOrder() throws Exception
	{	
		report1 = new ExtentReports("D:\\Saleema\\NormalOrderCashreport.html");
	  	log = report1.startTest("Place normal Order for Cash");
	  	
		Login();
		
		///click place order
		Actions act = new Actions(w);
		act.moveToElement(w.findElement(By.linkText("Place Order"))).click().perform();	
		WebElement Sublink = w.findElement(By.linkText("Stocks"));
		Sublink.click();
					
		////select buy option
		w.findElement(By.xpath(".//input[@value='BUY']")).click();
		
		///select market exchange
		Select ME = new Select(w.findElement(By.id("market_exchange")));
		ME.selectByVisibleText("NSE");
		
				
		///select stock name
		w.findElement(By.id("ssscrip")).sendKeys("dlf");
		w.findElement(By.partialLinkText("DLF LIMITED")).click();
				
		//qty
		w.findElement(By.id("stk_qty")).sendKeys("1");
				
		//price
		//w.findElement(By.id("stk_price")).clear();
		
		//w.findElement(By.id("stk_price")).sendKeys("250");
								
		///place order
		w.findElement(By.partialLinkText("PLACE ORDER")).click();
				
		///alert
		w.switchTo().alert().accept();

		///get text
		stckno = w.findElement(By.xpath("//*[@id='container']/div/p/span")).getText();
		if(w.getPageSource().contains(stckno))
		{
			System.out.println(stckno);
			
		}
		else
		{
			System.err.println("Error");
		}
						
	}
	
	@Test(priority = 7)
	public void ManageNormalBuyOrder() throws Exception
	{
		log = report1.startTest("Change Cash Order");
		
		String amocheck = "Error";
	  	if(w.getPageSource().contains(amocheck))
	  	{
	  		String capturemsg = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
	  		log.log(LogStatus.INFO, "Error as " +capturemsg);
	  	}
	  	else
	  	{
	  	///click watchlist>>place order
			Actions act = new Actions(w);
			act.moveToElement(w.findElement(By.linkText("Reports"))).click().perform();	
			WebElement Sublink = w.findElement(By.linkText("Order Status"));
			Sublink.click();
			
				///order status
				//w.findElement(By.partialLinkText("Order Status")).click();
				System.out.println("Stock no "+stckno);
				if(stckno !=null &&  !stckno.isEmpty())
				{
				String status = w.findElement(By.xpath("//div[@id='container']//td[contains(text(), 'OPN')]")).getText();
				System.out.println("status is " +status);
				String Status1 = "OPN";				
				System.out.println("Order no" +stckno);
				
				//check stkno
				w.findElement(By.xpath(".//input[@value='"+stckno+"|N']")).click();
				
				//JavascriptExecutor js = (JavascriptExecutor) w;
				//js.executeScript("0, 750");
				
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
					
					w.findElement(By.id("stk_qty")).sendKeys("5");
					
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
	  	}
				else
				{
					System.out.println("Order no not generated");
				}
	  	}
						
	}
		
	@Test(priority = 8)
	public void CancelNormalBuyOrder() throws Exception
	{				
		///cancel order
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
	}
	/*
	@Test(priority = 9)
	public void PlaceNormalSellOrder() throws Exception
	{
				///click place order
				Actions act = new Actions(w);
				act.moveToElement(w.findElement(By.linkText("Place Order"))).click().perform();	
				WebElement Sublink = w.findElement(By.linkText("Stocks"));
				Sublink.click();
									
				////select buy option
				w.findElement(By.xpath(".//input[@value='SELL']")).click();
				
				///select market exchange
				Select ME = new Select(w.findElement(By.id("market_exchange")));
				ME.selectByVisibleText("NSE");
				
						
				///select stock name
				w.findElement(By.id("ssscrip")).sendKeys("dlf");
				
				w.findElement(By.partialLinkText("DLF LIMITED")).click();
				//w.findElement(By.partialLinkText("ACC LIMITED")).click();
						
				//qty
				w.findElement(By.id("stk_qty")).sendKeys("2");
						
				//price
				//w.findElement(By.id("stk_price")).clear();
				
				//w.findElement(By.id("stk_price")).sendKeys("250");
						
				///place order
				w.findElement(By.partialLinkText("PLACE ORDER")).click();
						
				///alert
				w.switchTo().alert().accept();
				///get text
				String stckno = w.findElement(By.xpath("//*[@id='container']/div/p/span")).getText();
				
				System.out.println(stckno);
	}
	
	@Test(priority = 10)
	public void ManageNormalSellOrder() throws Exception
	{
		ManageNormalBuyOrder();
	}
	
	@Test(priority = 11)
	public void CancelNormalSellOrder() throws Exception
	{
		CancelNormalBuyOrder();
	}
	*/
	@Test(priority = 21000)
	  public void nouse()
	  {
		  report1.endTest(log);
		  report1.flush();
		  w.get("D:\\Saleema\\NormalOrderCashreport.html");
	  }
}
