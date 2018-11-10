///blank links are for logo, aadhar update same link name, refresh links equity and m/fbonds,

package com.TestCases;

import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import Main.BSOrderXL;

import org.testng.annotations.BeforeTest;

import java.io.FileInputStream;
import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;

public class TC8AllLinksTest extends BSOrderXL
{
	public static String url;
	public int count;
	
  @Test(priority=0)
  public void f() throws MalformedURLException, IOException 
  {	  	
		///report creation
		report1 = new ExtentReports("D:\\Saleema\\alllinkstest.html");
		log = report1.startTest("Test All Broken Links");
	  	
		////call login function
	  	Login();
	  	
	  	String popup = "Login Failed";
		
		if(!w.getPageSource().contains(popup))
		{
			count = 0;
			//check total links available ///listing all links and storing in list form
			List<WebElement> links = w.findElements(By.tagName("a"));		
			System.out.println("Total links present:" +links.size());
			log.log(LogStatus.INFO, "Total links present:" +links.size());
			{
				HttpURLConnection con=null;
				for(WebElement element:links)
				{
			        System.out.println("Link name " +element.getAttribute("innerText"));
			        
					log.log(LogStatus.INFO, "Working links :  " +element.getAttribute("innerText")); ///working bt display in loop
					//log.log(LogStatus.INFO, "Sub Menu's :" +element.getAttribute("innerText")); ///working bt display in loop
										
					url = element.getAttribute("href"); ///get href of tag a and store as url
					if(url!=null && !url.contains("javascript")) //to chk if url is not null and does not contain javascript
					{ 		//////it makes http call to server and response is returned
						con = (HttpURLConnection) (new URL(url).openConnection());
						con.connect();
						con.setConnectTimeout(3000);

						int respcode = con.getResponseCode();
						if(respcode==200)
						{					
							System.out.println("Valid links for URL " +url+ " is " +respcode+ "---" +con.getResponseMessage());
						}
						else
						{
							System.err.println("Broken links for URL " +url+ " is " +respcode+ "---" +con.getResponseMessage());
							log.log(LogStatus.INFO, "Broken links :  " +element.getAttribute("innerText")); ///working bt display in loop
							log.log(LogStatus.INFO, "Broken links for URL " +url+ " is " +respcode+ "---" +con.getResponseMessage());
						}
						con.disconnect();
					}
					count++;
				}
			}			
		}
		System.out.println("count total "+count);
		log.log(LogStatus.INFO, "Total links tested : " +count);
  }
    
  @AfterTest
	public void aftertest()
	{
		//w.close();
	}
		
	@Test(priority=40010)
	public void nouse()
	  {
		  report1.endTest(log);
		  report1.flush();
		  
		  w.findElement(By.partialLinkText("Logout")).click();
		  log.log(LogStatus.INFO, "Logged out Successfully");		
		  System.out.println("Application logged out");
			
		  w.get("D:\\Saleema\\alllinkstest.html");
	  }	
  

}
