package Trading;


import java.util.List;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.ArrayList;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
 
public class egtest {
 
  public static List findAllLinks(WebDriver driver)
 
  {
 
  List<WebElement> elementList = new ArrayList();
 
  elementList = driver.findElements(By.tagName("a"));
 
  elementList.addAll(driver.findElements(By.tagName("img")));
 
  List finalList = new ArrayList(); ;
 
  for (WebElement element : elementList)
 
  {
 
  if(element.getAttribute("href") != null)
 
  {
 
  finalList.add(element);
 
  }	  
 
  }
 
  return finalList;
 
  }
 
public static String isLinkBroken(URL url) throws Exception
 
{
 
//url = new URL("https://uat.kotaksecurities.com/itrade/user/watch.exe?action=C");
 
String response = "";
 
HttpURLConnection connection = (HttpURLConnection) url.openConnection();
 
try
 
{
 
    connection.connect();
 
     response = connection.getResponseMessage();	        
 
    connection.disconnect();
 
    return response;
 
}
 
catch(Exception exp)
 
{
 
return exp.getMessage();
 
}  

}
 
public static void main(String[] args) throws Exception {
 
// TODO Auto-generated method stub
	System.setProperty("webdriver.chrome.driver", "D:\\Selenium software\\drivers\\chromedriver.exe");
	WebDriver ff = new ChromeDriver();
 
ff.get("https://uat.kotaksecurities.com/itrade/user/welcome.exe?action=chk_seckey_stat");
 
//ff.get("http://www.yahoo.com/");	    

ff.findElement(By.id("userid")).sendKeys("1Y040");
ff.findElement(By.id("pwd")).sendKeys("login1");
ff.findElement(By.id("otp")).sendKeys("1111");
ff.findElement(By.partialLinkText("Login")).click();
System.out.println("Logged in Successfully");
 
    List<WebElement> allImages = findAllLinks(ff);    
 
    System.out.println("Total number of elements found " +allImages.size());
 
    for( WebElement element : allImages){
 
    	try
 
    	{
 
        System.out.println("URL: " + element.getAttribute("href")+ " returned " + isLinkBroken(new URL(element.getAttribute("href"))));
 
    	//System.out.println("URL: " + element.getAttribute("outerhtml")+ " returned " + isLinkBroken(new URL(element.getAttribute("href"))));
 
    	}
 
    	catch(Exception exp)
 
    	{
 
    	System.out.println("At " + element.getAttribute("innerHTML") + " Exception occured -&gt; " + exp.getMessage());	    
 
    	}
 
    }
 
    }
 
}