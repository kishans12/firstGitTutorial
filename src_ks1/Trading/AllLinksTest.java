package Trading;

import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import org.testng.annotations.BeforeTest;

import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;

public class AllLinksTest 
{
	ExtentReports report;
	ExtentTest log;
	WebDriver w;
	
  @Test
  public void f() throws MalformedURLException, IOException 
  {
	  String popup = "Login Failed";
		
		if(!w.getPageSource().contains(popup))
		{
		System.out.println("Logged in Successfully");
		//check total links available
		List<WebElement> links = w.findElements(By.tagName("a"));
		System.out.println("Total links present:" +links.size());
		log.log(LogStatus.INFO, "Total links present:" +links.size());
		//for(int i=0;i<links.size();i++)
							
		HttpURLConnection con=null;
		//HttpPost hp = new HttpPost("https://uat.kotaksecurities.com/itrade/user/watch.exe?show=menu");
		for(WebElement element:links)
		{
			String url = element.getAttribute("href");
			if(url!=null && !url.contains("javascript"))
			{
				con = (HttpURLConnection) (new URL(url).openConnection());
				con.connect();
				con.setConnectTimeout(3000);
				
				int respcode = con.getResponseCode();
				if(respcode==200)
				{
					//System.out.println("link name:" +links.get(i).getText());
					System.out.println("Valid links for URL " +url+ " is " +respcode+ "---" +con.getResponseMessage());	
				}
				else
				{
					System.err.println("Broken links for URL " +url+ " is " +respcode+ "---" +con.getResponseMessage());
					log.log(LogStatus.INFO, "Broken links for URL " +url+ " is " +respcode+ "---" +con.getResponseMessage());
				}
				//HttpPost hp = new HttpPost(url);
				//hp.releaseConnection();
			}
		}
		}		
			
  }
  
  @BeforeTest
  public void beforeTest() 
  {
	  	System.setProperty("webdriver.chrome.driver", "D:\\Selenium software\\chromedriver.exe");
		w = new ChromeDriver();
		w.manage().window().maximize();
		report = new ExtentReports("D:\\Saleema\\KotakSecurities\\src\\Trading\\AllLinksTest.html");
		log = report.startTest("Test All Broken Links");
		
		w.get("https://uat.kotaksecurities.com/itrade/user/welcome.exe?action=chk_seckey_stat");
		log.log(LogStatus.INFO, "Enter URL");
		//w.get("http://10.6.202.176/ctrade/user/cwelcome.exe");
				
						
		w.findElement(By.id("userid")).sendKeys("1Y302");
		log.log(LogStatus.INFO, "Enter username");
		w.findElement(By.id("pwd")).sendKeys("login1");
		log.log(LogStatus.INFO, "Enter Password");
		w.findElement(By.id("otp")).sendKeys("1111");
		log.log(LogStatus.INFO, "Enter security key");
		//w.findElement(By.partialLinkText("Login")).click();
		w.findElement(By.xpath("//*[@id='tbdiv']/span")).click();
		log.log(LogStatus.INFO, "Logged in Successfully");
  }

  @AfterTest
  public void afterTest() 
  {
	  	report.endTest(log);
		report.flush();
		
		w.findElement(By.partialLinkText("Logout")).click();
		log.log(LogStatus.INFO, "Logged out Successfully");		
		System.out.println("Application logged out");
	
		w.close();
  }

}
