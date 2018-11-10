package KeatPro;

import java.io.IOException;
import java.net.URL;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.winium.DesktopOptions;
import org.openqa.selenium.winium.WiniumDriver;
import org.testng.annotations.Test;

public class fortesting 
	{
		 @Test
		 public void test() throws IOException
		 {
		  DesktopOptions options= new DesktopOptions();
		  options.setApplicationPath("C:\\WINDOWS\\system32\\notepad.exe");
		  
		  //options.setApplicationPath("C:\\WINDOWS\\system32\\calc.exe");
		  WiniumDriver driver=new WiniumDriver(new URL("http://localhost:9999"),options);
		   
		   driver.findElementByName("Maximize").click();		   			
		   driver.findElementByClassName("Edit").sendKeys("This is sample test");
		   driver.findElementByName("File").click();
		   driver.findElementByName("Save").click();
		   driver.findElementByName("File name:").clear();
		   driver.findElementByName("File name:").sendKeys("abcdtest");
		   driver.findElementByName("Save").click();		   
		   driver.close();
		   System.out.println("Completed successfully");
		  
		 }
		}

//driver.findElementByClassName("4").click();
//driver.findElementByClassName("+").click();
//driver.findElementByClassName("8").click();
//driver.findElementByClassName("=").click();
//String get = driver.findElementById("158").getText();
//System.out.println("output "+get);


/*
String scrip2 = "Dear CLIENT_1Y040 Welcome to KEATProX ";
			  for(int i=0;i<5;i++)
			  {				  				  
				  WebElement scrip = w.findElementByClassName("SysListView32");
				  //scrip.findElement(By.name("Dear CLIENT_1Y040 Welcome to KEATProX Ver 2.2")).click();///working
////			  scrip.findElement(By.xpath("//*[contains(@Name, 'Dear CLIENT_1Y040 Welcome to KEATProX ')]")).click();///working
				  //scrip.findElement(By.xpath("//*[contains(@Name, '"+scrip2+"')]")).click();
				  String getmsg = scrip.findElement(By.xpath("//*[contains(@Name, '"+scrip2+"')]")).getAttribute("Name");///working
				  System.out.println("get message "+getmsg);
				  
			  }
			  
*/
