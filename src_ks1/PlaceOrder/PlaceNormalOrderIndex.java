package PlaceOrder;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.Test;

import Main.BSOrderXL;

public class PlaceNormalOrderIndex extends BSOrderXL
{
		public String stckno;
		public String Status1;
		public String status;
		public String stckno1;
		
		@Test(priority = 6)
		public void PlaceNormalBuyOrder() throws Exception
		{	
			w.findElement(By.id("userid")).sendKeys("1Y040");
			w.findElement(By.id("pwd")).sendKeys("login1");
			w.findElement(By.id("otp")).sendKeys("1111");
			w.findElement(By.partialLinkText("Login")).click();
			System.out.println("Logged in Successfully");
			
			///click place order
			Actions act = new Actions(w);
			act.moveToElement(w.findElement(By.linkText("Place Order"))).click().perform();	
			WebElement Sublink = w.findElement(By.linkText("Index F&O"));
			Sublink.click();
						
			////select buy option
			w.findElement(By.xpath(".//input[@value='BUY']")).click();
			
			///select market exchange
			Select ME = new Select(w.findElement(By.id("market_exchange")));
			ME.selectByVisibleText("NSE-Derivative");
			
					
			///select stock name
			w.findElement(By.id("ssscrip")).sendKeys("nif");
			w.findElement(By.partialLinkText("NIFTY")).click();
					
			///select expiry date
			Select ED = new Select(w.findElement(By.id("ssdrvexp")));
			ED.selectByVisibleText("26APR18");
			
			///select instrument type
			Select IT = new Select(w.findElement(By.id("ssdrvit")));
			IT.selectByVisibleText("FI");
			
			//enter no of lots
			w.findElement(By.id("stk_lot")).sendKeys("5");
					
			//price
			//w.findElement(By.id("stk_price")).clear();
			
			//w.findElement(By.id("stk_price")).sendKeys("250");
					
			///place order
			w.findElement(By.partialLinkText("PLACE ORDER")).click();
					
			///alert
			w.switchTo().alert().accept();
			
			///get text//*[@id="container"]/div/p/span[1]
			//stckno = w.findElement(By.xpath("//*[@id='container']/div/p/span")).getText();
			String stat = "ACTION NOT ALLOWED IN CURRENT PHASE";
			//stckno = w.findElement(By.xpath("//*[@id='container']/div/p/span")).getText();
			//System.out.println(stckno);
			if(w.getPageSource().contains(stat))
			{
				System.err.println("Error displayed as " +stat);
				PlaceNormalSellOrder();
				//System.out.println(stckno);
				
			}
			else
			{
				stckno = w.findElement(By.xpath("//*[@id='container']/div/p/span")).getText();
				System.out.println(stckno);
				//System.err.println("Error");
			}
								
		}
		
		@Test(priority = 7)
		public void ManageNormalBuyOrder() throws Exception
		{
					//stckno1 = stckno;
					///order status
					System.out.println("from managebuy order "+stckno+ " ....");
					w.findElement(By.partialLinkText("Order Status")).click();
						
					String status = w.findElement(By.xpath("//div[@id='container']//td[contains(text(), 'OPN')]")).getText();
					System.out.println("status is " +status);
					String Status1 = "OPN";				
					System.out.println("Order no " +stckno);
					
					//check stkno
					w.findElement(By.xpath(".//input[@value='"+stckno+"|Y']")).click();
					
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
						w.findElement(By.id("stk_lot")).clear();
						
						w.findElement(By.id("stk_lot")).sendKeys("2");
						
						///click on chnge order
						w.findElement(By.partialLinkText("CHANGE ORDER")).click();
										
						//cnfrm
						w.findElement(By.partialLinkText("Confirm")).click();
						System.out.println("Order no " +stckno+ " is modified successfully");
						///order status
						w.findElement(By.partialLinkText("Order Status")).click();
						
						//check stkno
						w.findElement(By.xpath(".//input[@value='"+stckno+"|Y']")).click();
					}
							
		}
			
		@Test(priority = 8)
		public void CancelNormalBuyOrder() throws Exception
		{			
			//stckno1 = stckno;
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
		
		@Test(priority = 9)
		public void PlaceNormalSellOrder() throws Exception
		{
					///click place order
					Actions act = new Actions(w);
					act.moveToElement(w.findElement(By.linkText("Place Order"))).click().perform();	
					WebElement Sublink = w.findElement(By.linkText("Index F&O"));
					Sublink.click();
										
					////select buy option
					w.findElement(By.xpath(".//input[@value='SELL']")).click();
					
					///select market exchange
					Select ME = new Select(w.findElement(By.id("market_exchange")));
					ME.selectByVisibleText("NSE-Derivative");
					
							
					///select stock name
					w.findElement(By.id("ssscrip")).sendKeys("f");
					w.findElement(By.partialLinkText("FTSE100")).click();
							
					///select expiry date
					Select ED = new Select(w.findElement(By.id("ssdrvexp")));
					ED.selectByVisibleText("20APR18");
					
					///select instrument type
					Select IT = new Select(w.findElement(By.id("ssdrvit")));
					IT.selectByVisibleText("FI");
					
					//enter no of lots
					w.findElement(By.id("stk_lot")).sendKeys("4");
							
					//price
					//w.findElement(By.id("stk_price")).clear();
					
					//w.findElement(By.id("stk_price")).sendKeys("250");
							
					///place order
					w.findElement(By.partialLinkText("PLACE ORDER")).click();
							
					///alert
					w.switchTo().alert().accept();
					///get text//*[@id="container"]/div/p/span[1]
					String stckno = w.findElement(By.xpath("//*[@id='container']/div/p/span")).getText();
					System.out.println("from sell order "+stckno+ " ....");
		}
		/*
		@Test(priority = 10)
		public void ManageNormalSellOrder() throws Exception
		{
			System.out.println("from managesell order "+stckno+ " ....");
			ManageNormalBuyOrder();
		}
		
		@Test(priority = 11)
		public void CancelNormalSellOrder() throws Exception
		{
			CancelNormalBuyOrder();
		}
		*/
}
