package KotakSec;

import org.testng.annotations.Test;

import io.appium.java_client.AppiumDriver;
import io.appium.java_client.MobileElement;
import io.appium.java_client.android.AndroidDriver;

import org.testng.annotations.BeforeTest;

import java.net.URL;

import org.openqa.selenium.By;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.testng.annotations.AfterTest;

public class Appfirstscrip 
{	
	DesiredCapabilities caps;
	AppiumDriver<MobileElement> driver;
	
  @Test
  public void f() throws Exception 
  {
	  //try
		//{
				driver = new AndroidDriver<MobileElement>(new URL("http://0.0.0.0:4723/wd/hub"), caps);
				Thread.sleep(1000);
				driver.get("https://uat.kotaksecurities.com/itrade/user/welcome.exe?action=chk_seckey_stat");
				driver.findElement(By.id("")).sendKeys("uname");////uname id
				driver.findElement(By.id("")).sendKeys("Pwd");///pwd id
				driver.findElement(By.id("")).sendKeys("acccode");//acccode
				driver.findElement(By.id("")).click();//login btn
				driver.findElement(By.id("")).click();//log out
				System.out.println("Successfully done");				
		//} 
	  //catch (MalformedURLException e) 
		//{
			//System.out.println(e.getMessage());
		//}
  }
  
  @BeforeTest
  public void beforeTest() 
  {
	  		caps = new DesiredCapabilities();
	  		//Set the Desired Capabilities
			caps.setCapability("deviceName", "My Phone"); //give device name of 
			caps.setCapability("udid", "ENUL6303030010"); //Give Device ID of your mobile phone
			caps.setCapability("platformName", "Android");
			caps.setCapability("platformVersion", "6.0");
			//caps.setCapability(MobileCapabilityType.DEVICE_NAME, "My Phone"); //another type of declaration
			caps.setCapability("appPackage", "com.android.vending"); //app package
			caps.setCapability("appActivity", "com.google.android.finsky.activities.MainActivity"); //app activity
			caps.setCapability("noReset", "true");
			caps.setCapability("newCommandTimeout", 2000);
  }

  @AfterTest
  public void afterTest() 
  {
	  //driver.quit();	  
  }

}
