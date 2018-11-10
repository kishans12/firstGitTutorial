package Trading;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

import Main.Login;

public class PlaceChangeCancelSorder
{
	//static Login L1 = new Login();
	
	public static void main() throws Exception
	{
		System.setProperty("webdriver.chrome.driver", "D:\\Selenium software\\drivers\\chromedriver.exe");
		WebDriver w = new ChromeDriver();
		
		/// get URL
		
		w.get("https://uat.kotaksecurities.com/itrade/user/welcome.exe?action=chk_seckey_stat");
		w.manage().window().maximize();
		w.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		
		////call login script
		
		//L1.getdetails();
		
		///enter username
		w.findElement(By.id("userid")).sendKeys("1y040");
		
		
		///enter password
		w.findElement(By.id("pwd")).sendKeys("login1");
		
		
		///enter acccess code
		w.findElement(By.id("otp")).sendKeys("1111");
		
		
		///click on login
		w.findElement(By.partialLinkText("Login")).click();
			
		////call createorder script
		//fn.Createorder();
		
		///click place order
		Actions act = new Actions(w);
		act.moveToElement(w.findElement(By.linkText("Place Order"))).click().perform();	
		WebElement Sublink = w.findElement(By.linkText("Stocks"));
		Sublink.click();
		
			
		////select sell option
		w.findElement(By.xpath(".//input[@value='SELL']")).click();
		
		///select market exchange
		Select ME = new Select(w.findElement(By.id("market_exchange")));
		ME.selectByVisibleText("NSE");
		
				
		///select stock name
		w.findElement(By.id("ssscrip")).sendKeys("dlf");
		
		//w.findElement(By.cssSelector("DLF LIMITED")).click();
		w.findElement(By.partialLinkText("DLF LIMITED")).click();
				
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
		
		///order status
		w.findElement(By.partialLinkText("Order Status")).click();
				
		String stat = "CAN";
		//String stat= w.findElement(By.xpath("//*[@id="container"]/div/form/div[1]/table/tbody/tr[8]/td[14]")).getText();
		//*[@id="container"]/div/form/div[1]/table/tbody/tr[8]/td[14]
		//*[@id="container"]/div/form/div[1]/table/tbody/tr[7]/td[14]
		if(stat.equals("CAN"))
		{
			System.out.println("Status of" +stckno+ " is Cancelled");
		}
		else
		{
			System.out.println("Status of" +stckno+ " is not in Cancel state");
		}
		if(w.getPageSource().contains(stckno))
			{
				//check stkno
				w.findElement(By.xpath(".//input[@value='"+stckno+"|N']")).click();
		
				//JavascriptExecutor js = (JavascriptExecutor) w;
				//js.executeScript("0, 750");

				///click change order
				w.findElement(By.partialLinkText("CHANGE ORDER")).click();
								
				//update stock no
				w.findElement(By.id("stk_qty")).clear();
				
				w.findElement(By.id("stk_qty")).sendKeys("10");
				
				///click on chnge order
				w.findElement(By.partialLinkText("CHANGE ORDER")).click();
								
				//cnfrm
				w.findElement(By.partialLinkText("Confirm")).click();
								
				///click order status
				w.findElement(By.partialLinkText("Order Status")).click();
				
				//check stkno
				//w.findElement(By.className("ordno")).click();
				w.findElement(By.xpath(".//input[@value='"+stckno+"|N']")).click();
								
				///cancel order
				w.findElement(By.partialLinkText("CANCEL ORDER")).click();
								
				//cnfrm
				w.findElement(By.partialLinkText("Confirm")).click();
				
				
				String status = "CAN";
				if(status.equals("CAN"))
				{
					System.out.println("Status of" +stckno+ " is Cancelled");
				}
				
			}
	}
		public static void main(String[] args) throws Exception
		{
			System.setProperty("webdriver.chrome.driver", "D:\\Selenium software\\drivers\\chromedriver.exe");
			WebDriver w = new ChromeDriver();
			
			/// get URL
			
			w.get("https://uat.kotaksecurities.com/itrade/user/welcome.exe?action=chk_seckey_stat");
			w.manage().window().maximize();
			w.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
			
			////call login script
			
			//L1.getdetails();
			
			///enter username
			w.findElement(By.id("userid")).sendKeys("1y040");
			
			
			///enter password
			w.findElement(By.id("pwd")).sendKeys("login1");
			
			
			///enter acccess code
			w.findElement(By.id("otp")).sendKeys("1111");
			
			
			///click on login
			w.findElement(By.partialLinkText("Login")).click();
				
			////call createorder script
			//fn.Createorder();
			
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
			w.findElement(By.id("ssscrip")).sendKeys("acc");
			
			//w.findElement(By.cssSelector("DLF LIMITED")).click();
			w.findElement(By.partialLinkText("ACC LIMITED")).click();
					
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
			
			///order status
			w.findElement(By.partialLinkText("Order Status")).click();
					
			String stat = "CAN";
			//String stat= w.findElement(By.xpath("//*[@id="container"]/div/form/div[1]/table/tbody/tr[8]/td[14]")).getText();
			//*[@id="container"]/div/form/div[1]/table/tbody/tr[8]/td[14]
			//*[@id="container"]/div/form/div[1]/table/tbody/tr[7]/td[14]
			if(stat.equals("CAN"))
			{
				System.out.println("Status of" +stckno+ " is Cancelled");
			}
			else
			{
				System.out.println("Status of" +stckno+ " is not in Cancel state");
			}
			if(w.getPageSource().contains(stckno))
				{
					//check stkno
					w.findElement(By.xpath(".//input[@value='"+stckno+"|N']")).click();
			
					//JavascriptExecutor js = (JavascriptExecutor) w;
					//js.executeScript("0, 750");

					///click change order
					w.findElement(By.partialLinkText("CHANGE ORDER")).click();
									
					//update stock no
					w.findElement(By.id("stk_qty")).clear();
					
					w.findElement(By.id("stk_qty")).sendKeys("10");
					
					///click on chnge order
					w.findElement(By.partialLinkText("CHANGE ORDER")).click();
									
					//cnfrm
					w.findElement(By.partialLinkText("Confirm")).click();
									
					///click order status
					w.findElement(By.partialLinkText("Order Status")).click();
					
					//check stkno
					//w.findElement(By.className("ordno")).click();
					w.findElement(By.xpath(".//input[@value='"+stckno+"|N']")).click();
									
					///cancel order
					w.findElement(By.partialLinkText("CANCEL ORDER")).click();
									
					//cnfrm
					w.findElement(By.partialLinkText("Confirm")).click();
					
					/*
					String status = "CAN";
					if(status.equals("CAN"))
					{
						System.out.println("Status of" +stckno+ " is Cancelled");
					}
					*/
				}
		//logout
		w.findElement(By.partialLinkText("Logout")).click();
				
		//close browser
		w.close();
		System.out.println("Succeeded");

	}

}
