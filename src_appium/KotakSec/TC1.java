package KotakSec;

import java.net.MalformedURLException;
import java.net.URL;

import org.openqa.selenium.remote.DesiredCapabilities;
import io.appium.java_client.AppiumDriver;
import io.appium.java_client.MobileElement;
import io.appium.java_client.android.AndroidDriver;
import io.appium.java_client.remote.MobileCapabilityType;

public class TC1 
{
	
	public static void main(String[] args) throws Exception,MalformedURLException 
	{		
		//Set the Desired Capabilities
		DesiredCapabilities caps = new DesiredCapabilities();
		caps.setCapability("deviceName", "My Phone"); //give device name of 
		caps.setCapability("udid", "ENUL6303030010"); //Give Device ID of your mobile phone
		caps.setCapability("platformName", "Android");
		caps.setCapability("platformVersion", "6.0");
		//caps.setCapability(MobileCapabilityType.DEVICE_NAME, "My Phone"); //another type of declaration
		caps.setCapability("appPackage", "com.android.vending"); //app package
		caps.setCapability("appActivity", "com.google.android.finsky.activities.MainActivity"); //app activity
		caps.setCapability("noReset", "true");
		caps.setCapability("newCommandTimeout", 2000);
		
		//Instantiate Appium Driver
		try
		{
				AppiumDriver<MobileElement> driver = new AndroidDriver<MobileElement>(new URL("http://0.0.0.0:4723/wd/hub"), caps);
				Thread.sleep(1000);
				driver.get("http://appium.io/");
				driver.quit();
				System.out.println("Successfully done");
				
		} catch (MalformedURLException e) 
		{
			System.out.println(e.getMessage());
		}
	}	
}