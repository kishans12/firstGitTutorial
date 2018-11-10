package Trading;

import java.io.FileInputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.LogStatus;

import Main.BSOrderXL;

public class Utilization extends BSOrderXL
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
	GTC.TC7GTCUtilisation pmc = new GTC.TC7GTCUtilisation();
	
	@Test(priority=0)
	public void PlaceGTCOrder() throws Exception
	{
		FileInputStream f = new FileInputStream("D:\\Saleema\\KotakSecurities\\testdata.xls");
		HSSFWorkbook wb = new HSSFWorkbook(f);
		s = wb.getSheet("GTC");
		HSSFSheet s1 = wb.getSheet("Watchlist");
		
		///report creation
		//report1 = new ExtentReports("D:\\Saleema\\NSEUtilisation.html");
		//log = report1.startTest("Manage Watchlist");
		//log = report1.startTest("Place GTC Cash Order to check utilisation");  	
		
		////call login function
		//Login();
			  	
		///call manage watchlist
		//cm.manage();
		
		log = report1.startTest("Place GTC Cash Order to check utilisation");
		
		///get multiple value from excel for utilization caln
	  	multiple = (int) s.getRow(3).getCell(1).getNumericCellValue();  	
		System.out.println("data " +multiple);
				
		///call capture total function to capture utilisation value
		captot.capturetotal();
				
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
		//w.findElement(By.id("price1")).clear();
		price = Double.parseDouble(pr1);
		System.out.println("price "+price);
		//w.findElement(By.id("price1")).sendKeys(String.valueOf(price));
		log.log(LogStatus.INFO, "Actual price is " +pr1);
		log.log(LogStatus.INFO, "Modified price is " +price);
		RM = ((quantity * price ) / multiple);
		roundoff = (Math.round(RM*100.0)/100.0);
		log.log(LogStatus.INFO, "Required Margin is " +roundoff);
		System.out.println("roundoff " +roundoff);
		
		///select date
		pmc.date();
		
		///click on place order
		//w.findElement(By.partialLinkText("Place Order")).click();  [same name on page]
		w.findElement(By.xpath("//*[@id=\'container\']/div/form/table/tbody/tr/td[1]/a")).click();
		log.log(LogStatus.INFO, "Click on Place Order");
		
		//check for any error
		String error = "Error";
		if(w.getPageSource().contains(error))
		{
			String capturemsg = w.findElement(By.xpath("//*[@id=\\'container\\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
			System.err.println("Error displayed as " +capturemsg);
			log.log(LogStatus.ERROR, "" +capturemsg);
		}
		else
		{			///click on confirm
			w.findElement(By.partialLinkText("Confirm")).click();
			log.log(LogStatus.INFO, "Click on Confirm");
			
			////capture gtc order no
			gtcorderno = w.findElement(By.xpath("//*[@id=\'container\']/div/form/table[3]/tbody/tr[3]/td[3]")).getText();
			log.log(LogStatus.INFO, "GTC Order got placed successfully and Order No is " +gtcorderno);
			
		}	
		
		Thread.sleep(150000);
		
		Actions act1 = new Actions(w);
		act1.moveToElement(w.findElement(By.linkText("Reports"))).click().perform();	
		WebElement Sublink1 = w.findElement(By.linkText("GTC Status"));
		Sublink1.click();
			
			System.out.println("gtcorderno out "+gtcorderno);
			chkitms = w.findElement(By.xpath("//input[starts-with(@value,'"+gtcorderno+"')]/following-sibling::td[2]")).getText();
			System.out.println("chkitms place "+chkitms);
			
			//to check if order is active
			if(chkitms.equals("0"))
			{
				log.log(LogStatus.ERROR, "Utilization cannot be calculated as order is ACTIVE");
			}
			else
			{
				///click funds >>equity/derivatives
				Actions act2 = new Actions(w);
				act2.moveToElement(w.findElement(By.linkText("Funds"))).click().perform();	
				WebElement Sublink2 = w.findElement(By.linkText("Equity/Derivatives"));
				Sublink2.click();
				log.log(LogStatus.INFO, "Click on Funds >> Equity/Derivatives");
				
				//click margin utilised cash&FNO
				w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).click();
				
				///check scrip and capture amount
				String scrip1 = s.getRow(7).getCell(1).getStringCellValue();
				
				if(w.getPageSource().contains(scrip1))
				{
					///get util amount of specific scrip				
					String total = w.findElement(By.xpath("//td[text()='"+scrip1+"']/following-sibling::td[6]")).getText();
					tot = Double.parseDouble(total);
					System.out.println("cap amt " +tot);
					Double match = tot - captot.cap;
					System.out.println("match captot "+match);
					matchround = Math.round(match*100.0)/100.0;
					System.out.println("f amt" +matchround);
					log.log(LogStatus.INFO, "Margin calculated is " +matchround);
					
					////to check if reqd margin and captured util amt are same
					if(matchround.equals(roundoff))
					{
						///click on cancel
						w.findElement(By.partialLinkText("CANCEL")).click();
						
						///click on total utilised cashNfno
						String totalmargin = w.findElement(By.xpath("//*[@id=\'limit-node-1\']/td[3]")).getText();
						String t = totalmargin.replaceAll(",","").trim();
						String m = t.replaceAll("","");
						System.out.println(m);
						log.log(LogStatus.INFO, "Total Margin is " +totalmargin);
						
						///click on margin utilised cashNfno
						String marginutilised = w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).getText();
						String t1 = marginutilised.replaceAll(",","").trim();
						String m1 = t1.replaceAll("","");
						System.out.println(m1);
						log.log(LogStatus.INFO, "Margin Utilised is " +marginutilised);
						
						///click on margin available cashNfno
						String marginavailable = w.findElement(By.xpath("//*[@id=\'limit-node-6\']/td[3]/strong")).getText();
						String t2 = marginavailable.replaceAll(",","").trim();
						String m2 = t2.replaceAll("","");
						System.out.println(m2);
						log.log(LogStatus.INFO, "Margin Available is " +marginavailable);
						
						//verify totalmargin - margin utilsed = margin available
						Double ab = Double.parseDouble(m)-Double.parseDouble(m1);
						BigDecimal bd = new BigDecimal(ab);
						DecimalFormat df = new DecimalFormat("0.00");
						df.setMaximumFractionDigits(2);
						sd = df.format(ab);
						System.out.println(sd);
						
						///check if above are equal
						if(sd.equals(m2))
						{
							System.out.println("Equal");
							log.log(LogStatus.INFO, "Utilized -- As Required Margin and Margin Utilised are Equal");
							log.log(LogStatus.INFO, "Total Margin & Margin utilized are equal to Margin Available " +sd);
						}
						
					}
					else
					{
						log.log(LogStatus.ERROR, "Not Utilized -- As Required Margin and Margin Utilised are Unequal");
						log.log(LogStatus.ERROR, "Total Margin & Margin utilized are not equal to Margin Available ");
					}
				}
			}
	}
					
	@Test(priority=1)
	public void CancelGTCOrder() throws Exception
	{
		log = report1.startTest("Cancel GTC Cash Order of utilisation");
							
		///check for any error
		String amocheck = "Error";
		if(w.getPageSource().contains(amocheck))
		{
			String capturemsg = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
		  	log.log(LogStatus.ERROR, "" +capturemsg);
		}
		else
		{
			///click watchlist>>place order
			Actions act = new Actions(w);
			act.moveToElement(w.findElement(By.linkText("Reports"))).click().perform();	
			WebElement Sublink = w.findElement(By.linkText("GTC Status"));
			Sublink.click();
			log.log(LogStatus.INFO, "Click on Reports >> GTC Status");
							
			log.log(LogStatus.INFO, "Select GTC order No " +gtcorderno);
			
			if(chkitms.equals("0"))
			{
				//click cancel
				w.findElement(By.xpath("//input[starts-with(@value,'"+gtcorderno+"')]/following-sibling::td[13]/a/img")).click();
				log.log(LogStatus.INFO, "Click on Cancel");
				System.out.println("chkitms 0 "+chkitms);
												
				String err = "Error";
				if(w.getPageSource().contains(err))
				{
					String capturemsg = w.findElement(By.xpath("//*[@id=\'container\']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
					log.log(LogStatus.ERROR, "" +capturemsg);
					w.findElement(By.partialLinkText("BACK")).click();
				}
				
				//click confirm
				w.findElement(By.partialLinkText("Confirm")).click();
				log.log(LogStatus.INFO, "Click on Confirm");
				
				log.log(LogStatus.INFO, "Order No " +gtcorderno+ " is cancelled successfully");
			}
				else
				{
					//click cancel
					w.findElement(By.xpath("//input[starts-with(@value,'"+gtcorderno+"')]/following-sibling::td[12]/a/img")).click();
					log.log(LogStatus.INFO, "Click on Cancel");
					System.out.println("chkitms no 0 "+chkitms);
					
					String err = "Error";
					if(w.getPageSource().contains(err))
					{
						String capturemsg = w.findElement(By.xpath("//*[@id=\'container\']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
						log.log(LogStatus.ERROR, "" +capturemsg);
						w.findElement(By.partialLinkText("BACK")).click();
					}
					
					//click confirm
					w.findElement(By.partialLinkText("Confirm")).click();
					log.log(LogStatus.INFO, "Click on Confirm");
					
					log.log(LogStatus.INFO, "Order No " +gtcorderno+ " is cancelled successfully");
					
					Double newamt = captot.cap;
					System.out.println("newamt "+newamt);
					///call capture total
					captot.capturetotal();
					Double canamt = newamt - captot.cap;
					System.out.println("can amt " +canamt);
					log.log(LogStatus.INFO, "Cancelled Required Margin is " +canamt);
						
				}
			
		}
	}
	
	@Test(priority=36)
	public void nouse()
	  {
		 report1.endTest(log);
		 report1.flush();
		 w.get("D:\\Saleema\\GTCUtilisation.html");
	  }
}
