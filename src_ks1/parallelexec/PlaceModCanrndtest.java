package parallelexec;

import java.io.FileInputStream;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

public class PlaceModCanrndtest
{
	public static String gtcorderno;
	public static HSSFSheet s;
	public static Double roundoff;
	public static Double RM;
	public static int multiple;
	public static Double price;
	public static Double matchround;
	public static Double tot;
	public static String sd;
	public static String chkitms;
	
	Utilisation.NSEUtilisation captot = new Utilisation.NSEUtilisation();
	
	@Test
	public void PlaceGTCOrder() throws Exception
	{
		  WebDriver w1;
		  ExtentReports report1;
		  ExtentTest log;
		  String Uname = "1Y040";
		  String Pwd = "login1";
		  String AccCode = "1111";
		  
		System.setProperty("webdriver.chrome.driver", "D:\\Selenium software\\chromedriver.exe");
		w1 = new ChromeDriver();
		
		/// get URL
		
		w1.get("https://uat.kotaksecurities.com/itrade/user/welcome.exe?action=chk_seckey_stat");
		
		w1.manage().window().maximize();
		w1.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
						
		///report creation
		report1 = new ExtentReports("D:\\Saleema\\exec1.html");
		//log = report1.startTest("Manage Watchlist");
		log = report1.startTest("Place GTC Cash Order");  	
		
		////call login function
		w1.findElement(By.id("userid")).sendKeys(Uname);
	  	log.log(LogStatus.INFO, "Enter User name " +Uname);
		w1.findElement(By.id("pwd")).sendKeys(Pwd);
		log.log(LogStatus.INFO, "Enter Password ******");
		w1.findElement(By.id("otp")).sendKeys(AccCode);
		log.log(LogStatus.INFO, "Enter Security Key/Access Code ****");
		w1.findElement(By.partialLinkText("Login")).click();
		System.out.println("Logged in Successfully");
		log.log(LogStatus.INFO, "Logged in Successfully");
		 	
				
		///click watchlist>>place order
		Actions act = new Actions(w1);
		act.moveToElement(w1.findElement(By.linkText("Watchlist"))).click().perform();	
		WebElement Sublink = w1.findElement(By.linkText("My Watchlists"));
		Sublink.click();
		log.log(LogStatus.INFO, "Click on Watchlist >> My Watchlists");
		
				
		report1.endTest(log);
		report1.flush();
		w1.get("D:\\Saleema\\exec1.html");
	}
	
	@Test
	public void PlaceGTCOrder1() throws Exception
	{
		   WebDriver w2;
		   ExtentReports report2;
		   ExtentTest log;
		   String Uname = "1Y040";
		   String Pwd = "login1";
		   String AccCode = "1111";
		  
		   System.setProperty("webdriver.chrome.driver", "D:\\Selenium software\\chromedriver.exe");
			w2 = new ChromeDriver();
			
			/// get URL
			
			w2.get("https://uat.kotaksecurities.com/itrade/user/welcome.exe?action=chk_seckey_stat");
			
			w2.manage().window().maximize();
			w2.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
			
		///report creation
		report2 = new ExtentReports("D:\\Saleema\\exec2.html");
		//log = report1.startTest("Manage Watchlist");
		log = report2.startTest("Place GTC Cash Order");  	
		
		////call login function
		w2.findElement(By.id("userid")).sendKeys(Uname);
	  	log.log(LogStatus.INFO, "Enter User name " +Uname);
		w2.findElement(By.id("pwd")).sendKeys(Pwd);
		log.log(LogStatus.INFO, "Enter Password ******");
		w2.findElement(By.id("otp")).sendKeys(AccCode);
		log.log(LogStatus.INFO, "Enter Security Key/Access Code ****");
		w2.findElement(By.partialLinkText("Login")).click();
		System.out.println("Logged in Successfully");
		log.log(LogStatus.INFO, "Logged in Successfully");
				
		///click funds >>equity/derivatives
				Actions act11 = new Actions(w2);
				act11.moveToElement(w2.findElement(By.linkText("Funds"))).click().perform();	
				WebElement Sublink11 = w2.findElement(By.linkText("Equity/Derivatives"));
			Sublink11.click();
					
			///click watchlist>>place order
			Actions act = new Actions(w2);
			act.moveToElement(w2.findElement(By.linkText("Watchlist"))).click().perform();	
			WebElement Sublink = w2.findElement(By.linkText("My Watchlists"));
			Sublink.click();
			log.log(LogStatus.INFO, "Click on Watchlist >> My Watchlists");
			
		report2.endTest(log);
		report2.flush();
		w2.get("D:\\Saleema\\exec2.html");
	}
	
	
}
