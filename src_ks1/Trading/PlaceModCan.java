package Trading;

import java.io.FileInputStream;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.LogStatus;

import Main.BSOrderXL;

public class PlaceModCan extends BSOrderXL
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
	//GTC.Utilization gtcutil = new GTC.Utilization();
	
	public void date() throws Exception
	{
		String window = w.getWindowHandle();
		w.findElement(By.xpath("//*[@id=\'gtctable\']/tbody/tr[2]/td[8]/a/img")).click();
		System.out.println(w.getTitle());
		for(String winHandle : w.getWindowHandles())
		{
		    w.switchTo().window(winHandle);
		}
		System.out.println(w.getTitle());
		
		//to print month
		Date today = Calendar.getInstance().getTime();
		SimpleDateFormat formatter = new SimpleDateFormat("MM");
		String month1 = formatter.format(today);
		
		//to print year
		SimpleDateFormat year = new SimpleDateFormat("yyyy");
		String yr = year.format(today);
		
		//to print day
		DateFormat dateFormat = new SimpleDateFormat("d");
		Calendar cal = Calendar.getInstance();
		cal.setTime(new Date());
		cal.add(Calendar.DATE, 1);
		String newDate = dateFormat.format(cal.getTime());
		System.out.println("newdate "+newDate);
		//click on next day
		w.findElement(By.xpath("//a[text()= '"+newDate+"']")).click();
		w.switchTo().window(window);
		System.out.println("back to window");
		log.log(LogStatus.INFO, "Selected Date is "+newDate+ "/" +month1+ "/" +yr);
	}
	
	@Test(priority=0)
	public void PlaceGTCOrder() throws Exception
	{
		FileInputStream f = new FileInputStream("D:\\Saleema\\KotakSecurities\\testdata.xls");
		HSSFWorkbook wb = new HSSFWorkbook(f);
		s = wb.getSheet("GTC");
		HSSFSheet s1 = wb.getSheet("Watchlist");
		
		///report creation
		//report1 = new ExtentReports("D:\\Saleema\\GTC.html");
		//log = report1.startTest("Manage Watchlist");
		//log = report1.startTest("Place GTC Cash Order");  	
		
		////call login function
		//Login();
			  	
		///call manage watchlist
		//cm.manage();
		
		log = report1.startTest("Place GTC Cash Order");
		
		///get multiple value from excel for utilization caln
	  	multiple = (int) s.getRow(3).getCell(1).getNumericCellValue();  	
		System.out.println("data " +multiple);
				
		///call capture total function to capture utilisation value
		//captot.capturetotal();
				
		///click watchlist>>place order
		Actions act = new Actions(w);
		act.moveToElement(w.findElement(By.linkText("Watchlist"))).click().perform();	
		WebElement Sublink = w.findElement(By.linkText("My Watchlists"));
		Sublink.click();
		log.log(LogStatus.INFO, "Click on Watchlist >> My Watchlists");
		
		///watchlistname
		String watchlistname = s1.getRow(1).getCell(1).getStringCellValue();
		w.findElement(By.xpath("//div[@id='container']//a[contains(text(), '"+watchlistname+"')]")).click();
		
		///select scrip
		String xcelscrip = s.getRow(5).getCell(1).getStringCellValue();
		System.out.println("data1 " +xcelscrip);
		w.findElement(By.xpath(".//input[@value='"+xcelscrip+"']")).click();
						
		///click GTC
		w.findElement(By.partialLinkText("GTC ORDER")).click();
		log.log(LogStatus.INFO, "Click on GTC Order");
		
		///check checkbox
		w.findElement(By.id("order1")).click();
		log.log(LogStatus.INFO, "Select Checkbox");
		
		///get scripname
		String scripname = w.findElement(By.id("name1")).getAttribute("value");
		log.log(LogStatus.INFO, "Selected Scrip is " +scripname);
				
		///Select buy/sell value
		w.findElement(By.id("bsind1")).click();
		String value = s.getRow(6).getCell(1).getStringCellValue();
		Select val = new Select(w.findElement(By.id("bsind1")));
		val.selectByVisibleText(value);
		log.log(LogStatus.INFO, "Selected Value is " +value);
		
		////enter quantity
		int quantity = (int) s.getRow(4).getCell(1).getNumericCellValue();  	
		w.findElement(By.id("qty1")).sendKeys(String.valueOf(quantity));
		log.log(LogStatus.INFO, "Entered Quantity is: " +quantity);
		
		///enter price
		String pr1 = w.findElement(By.id("price1")).getAttribute("value");
		w.findElement(By.id("price1")).clear();
		price = Double.parseDouble(pr1)-10;
		System.out.println("price "+price);
		w.findElement(By.id("price1")).sendKeys(String.valueOf(price));
		log.log(LogStatus.INFO, "Actual price is " +pr1);
		log.log(LogStatus.INFO, "Modified price is " +price);
				
		///select date
		date();
		
		///click on place order
		//w.findElement(By.partialLinkText("Place Order")).click();
		w.findElement(By.xpath("//*[@id='container']/div/form/table/tbody/tr/td[1]/a")).click();
		log.log(LogStatus.INFO, "Click on Place Order");
		
		String error = "Error";
		if(w.getPageSource().contains(error))
		{
			String capturemsg = w.findElement(By.xpath("//*[@id='container']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
			System.err.println("Error displayed as " +capturemsg);
			log.log(LogStatus.ERROR, "" +capturemsg);
		}
		else
		{			///click on confirm
			w.findElement(By.partialLinkText("Confirm")).click();
			log.log(LogStatus.INFO, "Click on Confirm");
			
			////capture gtc order no
			gtcorderno = w.findElement(By.xpath("//*[@id='container']/div/form/table[3]/tbody/tr[3]/td[3]")).getText();
			log.log(LogStatus.INFO, "GTC Order got placed successfully and Order No is " +gtcorderno);
		}	
	}
	
	@Test(priority=1)
	public void ManageGTCOrder() throws Exception
	{
		log = report1.startTest("Change GTC Cash Order");
		
		///click watchlist>>place order
		Actions act = new Actions(w);
		act.moveToElement(w.findElement(By.linkText("Reports"))).click().perform();	
		WebElement Sublink = w.findElement(By.linkText("GTC Status"));
		Sublink.click();
		log.log(LogStatus.INFO, "Click on Reports >> GTC Status");
		
		chkitms = w.findElement(By.xpath("//*[@id=\'tbdiv\']/table/tbody/tr[4]/td[2]/table/tbody/tr[2]/td[2]")).getText();
		System.out.println("chkitms manage "+chkitms);
		if(chkitms.equals("0"))
		{
			log.log(LogStatus.INFO, "Select GTC order No " +gtcorderno);
			
			///click on modify
			w.findElement(By.xpath("//input[starts-with(@value,'"+gtcorderno+"')]/following-sibling::td[12]/a/img")).click();
			log.log(LogStatus.INFO, "Click on Modify");
			
			String error = "Error";
			if(w.getPageSource().contains(error))
			{
				String capturemsg = w.findElement(By.xpath("//*[@id=\'container\']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
				System.err.println("Error displayed as " +capturemsg);
				log.log(LogStatus.ERROR, "" +capturemsg);
			}
			
			///modify quantity
			w.findElement(By.id("qty")).clear();
			int quantity = (int) s.getRow(4).getCell(6).getNumericCellValue();  	
			w.findElement(By.id("qty")).sendKeys(String.valueOf(quantity));
			log.log(LogStatus.INFO, "Modified Quantity : " +quantity);
			
			//click place order
			w.findElement(By.xpath("//*[@id=\'container\']/div/form/table[3]/tbody/tr/td[1]/a")).click();
			log.log(LogStatus.INFO, "Click on Place Order");
			
			//confirm
			w.findElement(By.partialLinkText("CONFIRM")).click();
			log.log(LogStatus.INFO, "Click on Confirm");
			
			log.log(LogStatus.INFO, "Order No " +gtcorderno+ " is modified successfully");
		}
				
	}
	
	@Test(priority=2)
	public void CancelGTCOrder() throws Exception
	{
		log = report1.startTest("Cancel GTC Cash Order");
		
		
		///check for any error
		String amocheck = "Error";
		if(w.getPageSource().contains(amocheck))
		{
			String capturemsg = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
		  	log.log(LogStatus.ERROR, "" +capturemsg);
		}
		else
		{
			///select status
			w.findElement(By.partialLinkText("Status")).click();
			log.log(LogStatus.INFO, "Click on Status");
			
			log.log(LogStatus.INFO, "Select GTC order No " +gtcorderno);
			
			//click cancel
			w.findElement(By.xpath("//input[starts-with(@value,'"+gtcorderno+"')]/following-sibling::td[13]/a/img")).click();
			log.log(LogStatus.INFO, "Click on Cancel");
			
			String err = "Error";
			if(w.getPageSource().contains(err))
			{
				String capturemsg = w.findElement(By.xpath("//*[@id=\'container\']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
				log.log(LogStatus.ERROR, "" +capturemsg);
				w.findElement(By.partialLinkText("BACK")).click();
			}
			else
			{
				//click confirm
				w.findElement(By.partialLinkText("Confirm")).click();
				log.log(LogStatus.INFO, "Click on Confirm");
				
				log.log(LogStatus.INFO, "Order No " +gtcorderno+ " is cancelled successfully");
			}
		}
	}
	
	@Test(priority=36)
	public void nouse()
	  {
		 //report1.endTest(log);
		 //report1.flush();
		 //w.get("D:\\Saleema\\GTCUtilisation.html");
	  }
}
