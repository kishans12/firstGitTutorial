package Utilisation;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.TimeZone;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import Main.BSOrderXL;

public class ManageWatchlist extends BSOrderXL
{
	public static HSSFSheet s;
	public static HSSFSheet s1;
	public static HSSFSheet sw;
	public static String watchlistname;
	public static String chk;
	public static WebElement ele;
	public static String NSE = "//*[@id=\'container\']/div/form/table[2]/tbody/tr[2]/td[1]/input";
	public static String BSE = "//*[@id=\'container\']/div/form/table[2]/tbody/tr[4]/td[1]/input";
	public static String Derivative = "//*[@id=\'container\']/div/form/table[2]/tbody/tr[2]/td[2]/input";
	public static String Currency = "//*[@id=\'container\']/div/form/table[2]/tbody/tr[2]/td[2]/input";
	public static Row rw;
	public static Cell cell;
	public static String scrip;
	public static String get;
	public static String dte;
	public static String cmprdate;
	public static String cmprmonth;
	public static String cmpryr;
	public static String a;
	public static String b;
	public static WebElement a1;
	public static String getexactexpdate;
	public static String combine;

	
	public void getdate()
	{
		List<WebElement> myList= w.findElements(By.xpath(".//td[input[@type = 'checkbox']]"));
		List<WebElement> myListele = w.findElements(By.xpath("//input[@name='stockcode' and @type = 'checkbox']"));
		System.out.println("Total elements :" +myList.size());

        Date currDate = Calendar.getInstance().getTime();
	   	SimpleDateFormat dateFormat = new SimpleDateFormat("ddMMMyy");
		SimpleDateFormat dateFormatformonth = new SimpleDateFormat("MMM");
		
		String getcurrdate = dateFormat.format(currDate);
		System.out.println("curr date "+getcurrdate);
	       		 
		Calendar calendar = Calendar.getInstance(TimeZone.getDefault());
		int year = calendar.get(Calendar.YEAR);
		int date = calendar.get(Calendar.DAY_OF_MONTH);
		String month = dateFormatformonth.format(currDate);
		dte = String.valueOf(date);
		
        String arr[]=new String[myList.size()];

 	    for(int i=0;i<myList.size();i++)
	 	{			
 	    	WebElement ele;
			arr[i]= myList.get(i).getText();		
			String actualscrip = arr[i];
 	    	ele = myListele.get(i);
 	    	ele.click();
			
			System.out.println("actualscrip "+actualscrip);
			String[] getscripexpirydate = actualscrip.split("]" ,0);
			String gsed = getscripexpirydate[1];
			System.out.println("gsed "+gsed);
			if(gsed.startsWith("["))
			{
				String[]removebrac = gsed.split("\\[");
				System.out.println("1 "+removebrac[1]);
				getexactexpdate = removebrac[1];
				        		
				String[] getonlydate = getexactexpdate.split("(?<=[A-Z])(?=[0-9])|(?<=[0-9])(?=[A-Z])");
				cmprdate = getonlydate[0];
				cmprmonth = getonlydate[1];
				
				int result = dte.compareTo(cmprdate);
				System.out.println("result "+result);
				if(result<=0)
				{
					if(month.toUpperCase().equals(cmprmonth))
        			{
						b = ele.getAttribute("value");
			  			System.out.println("getvalue " +b);
			  			w.findElement(By.partialLinkText("SUBMIT")).click();
			  			break;
        			}
					else
					{
						ele.click();
					}		
				}
				else
				{
					ele.click();
				}
			}
			else
			{				
				b = ele.getAttribute("value");
	  			System.out.println("getvalue " +b);
	  			w.findElement(By.partialLinkText("SUBMIT")).click();
				break;
			}
	 	}
	}
	
	@Test(priority=0)
	public void manage() throws Exception
	{
		//FileInputStream f = new FileInputStream(srcfile);
		//wb = new HSSFWorkbook(f);
		s1 = wb.getSheet("AddScrip");
		s = wb.getSheet("Data");
		sw = wb.getSheet("Result");
		
		///report creation
		//report1 = new ExtentReports("D:\\Saleema\\Manage.html");
		//log = report1.startTest("Manage Watchlist");
	  	
		//Login();
		
		///click funds >>equity/derivatives
		Actions act = new Actions(w);
		act.moveToElement(w.findElement(By.linkText("Watchlist"))).click().perform();	
		WebElement Sublink = w.findElement(By.linkText("Manage"));
		Sublink.click();
		log.log(LogStatus.INFO, "Click on Watchlist >> Manage");
		
		///click Add
		w.findElement(By.partialLinkText("ADD")).click();
		log.log(LogStatus.INFO, "Click on Add");
		
		//enter watchlist name
		watchlistname = s1.getRow(1).getCell(1).getStringCellValue();  	
		w.findElement(By.xpath("//*[@id=\'container\']/div/form/table[1]/tbody/tr/td[2]/input")).sendKeys(watchlistname);
		log.log(LogStatus.INFO, "Enter Watchlist Name : " +watchlistname);
		
		///click submit
		w.findElement(By.partialLinkText("Submit")).click();
		log.log(LogStatus.INFO, "Click on Submit");
		
		///to check if watchlistname already exist
		String check = "Error";
		if(w.getPageSource().contains(check))
		{
			String catcherror = w.findElement(By.xpath("//*[@id=\'container\']/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td")).getText();
			log.log(LogStatus.INFO, "" +catcherror);
			w.findElement(By.partialLinkText("BACK")).click();
			
			///go to manage
			Actions act1 = new Actions(w);
			act1.moveToElement(w.findElement(By.linkText("Watchlist"))).click().perform();	
			WebElement Sublink1 = w.findElement(By.linkText("Manage"));
			Sublink1.click();
			
			///check and delete same name
			if(w.getPageSource().contains(watchlistname))
			{
				List<WebElement> radio = w.findElements(By.xpath("//input[@name='id' and @type = 'radio']"));
				for(int i=0;i<radio.size();i++)
				{
					ele = radio.get(i);
					String value=ele.getAttribute("id");
					chk =w.findElement(By.xpath("//div[@id='container']//a[contains(text(), '"+watchlistname+"')]")).getText();
				}		
				if(chk.contains(watchlistname))
				{
					ele.click();
				}
				
				w.findElement(By.partialLinkText("DELETE")).click();
				log.log(LogStatus.INFO, "Existed Watchlist deleted successfully");
				w.findElement(By.partialLinkText("ADD")).click();
				log.log(LogStatus.INFO, "Add watchlist");
				w.findElement(By.xpath("//*[@id=\'container\']/div/form/table[1]/tbody/tr/td[2]/input")).sendKeys(watchlistname);
				log.log(LogStatus.INFO, "Enter Watchlist Name : " +watchlistname);
				w.findElement(By.partialLinkText("Submit")).click();
				log.log(LogStatus.INFO, "Click on Submit");
				w.findElement(By.xpath("//div[@id='container']//a[contains(text(), '"+watchlistname+"')]")).click();
				log.log(LogStatus.INFO, "Select " +watchlistname+ " to add scrips");
				
			}
		}
		else
			{
				///select added watchlist
				w.findElement(By.xpath("//div[@id='container']//a[contains(text(), '"+watchlistname+"')]")).click();
				log.log(LogStatus.INFO, "Watchlist " +watchlistname+ " gets added successfully " );
				
			}
		
		//addallscrips();
		addloopscrips();
		
	}
	
	public void addallscrips() throws Exception
	{
		FileOutputStream Of = new FileOutputStream(srcfile);

		///add nse scrip
		w.findElement(By.partialLinkText("ADD")).click();
		String nse = s1.getRow(5).getCell(1).getStringCellValue();
		w.findElement(By.id("search_list")).sendKeys(nse);
		w.findElement(By.id("me")).click();
		Select ME = new Select(w.findElement(By.id("me")));
		ME.selectByValue("NN");
		w.findElement(By.partialLinkText("SUBMIT")).click();
		//w.findElement(By.xpath("//div[@id='container']//input[contains(text(), '"+nse+"')]")).click();
		w.findElement(By.xpath(NSE)).click();
		String getnsevalue = w.findElement(By.xpath(NSE)).getAttribute("value");
		System.out.println("getvalue " +getnsevalue);
		
		rw = s.getRow(5);
		rw = s1.getRow(5);
		cell = rw.createCell(5);
		cell.setCellValue(getnsevalue);
		
		w.findElement(By.partialLinkText("SUBMIT")).click();
		log.log(LogStatus.INFO, "Scrip "+nse.toUpperCase()+ " got added for NSE");
						
		///add bse scrip
		w.findElement(By.partialLinkText("ADD")).click();
		String bse = s1.getRow(6).getCell(1).getStringCellValue();
		w.findElement(By.id("search_list")).sendKeys(bse);
		w.findElement(By.id("me")).click();
		Select ME1 = new Select(w.findElement(By.id("me")));
		ME1.selectByValue("BN");
		w.findElement(By.partialLinkText("SUBMIT")).click();
		//w.findElement(By.xpath("//div[@id='container']//a[contains(text(), '"+bse+"')]")).click();
		w.findElement(By.xpath(BSE)).click();
		String getbsevalue = w.findElement(By.xpath(BSE)).getAttribute("value");
		System.out.println("getvalue " +getbsevalue);
		
		rw = s.getRow(6);
		rw = s1.getRow(6);
		cell = rw.createCell(5);
		cell.setCellValue(getbsevalue);
		
		w.findElement(By.partialLinkText("SUBMIT")).click();
		log.log(LogStatus.INFO, "Scrip "+bse.toUpperCase()+ " got added for BSE");
		
		///add derivative scrip
		w.findElement(By.partialLinkText("ADD")).click();
		String derivative = s1.getRow(7).getCell(1).getStringCellValue();
		w.findElement(By.id("search_list")).sendKeys(derivative);
		w.findElement(By.id("me")).click();
		Select ME2 = new Select(w.findElement(By.id("me")));
		ME2.selectByValue("ND");
		w.findElement(By.partialLinkText("SUBMIT")).click();
		//w.findElement(By.xpath("//div[@id='container']//a[contains(text(), '"+derivative+"')]")).click();
		w.findElement(By.xpath(Derivative)).click();
		String getdervalue = w.findElement(By.xpath(Derivative)).getAttribute("value");
		System.out.println("getvalue " +getdervalue);
		
		rw = s.getRow(7);
		rw = s1.getRow(7);
		cell = rw.createCell(5);
		cell.setCellValue(getdervalue);
		
		w.findElement(By.partialLinkText("SUBMIT")).click();
		log.log(LogStatus.INFO, "Scrip "+derivative.toUpperCase()+ " got added for Derivatives");
		
		///add currency scrip
		w.findElement(By.partialLinkText("ADD")).click();
		String currency = s1.getRow(8).getCell(1).getStringCellValue();
		w.findElement(By.id("search_list")).sendKeys(currency);
		w.findElement(By.id("me")).click();
		Select ME3 = new Select(w.findElement(By.id("me")));
		ME3.selectByValue("CN");
		w.findElement(By.partialLinkText("SUBMIT")).click();
		//w.findElement(By.xpath("//div[@id='container']//a[contains(text(), '"+currency+"')]")).click();
		w.findElement(By.xpath(Currency)).click();
		String getcurrvalue = w.findElement(By.xpath(Currency)).getAttribute("value");
		System.out.println("getvalue " +getcurrvalue);
		
		rw = s.getRow(8);
		rw = s1.getRow(8);
		cell = rw.createCell(5);
		cell.setCellValue(getcurrvalue);
		
		w.findElement(By.partialLinkText("SUBMIT")).click();
		log.log(LogStatus.INFO, "Scrip "+currency.toUpperCase()+ " got added for Currency");
		log.log(LogStatus.PASS,"Passed");
		
		wb.write(Of);
		Of.close();
	}
	
	public void addloopscrips() throws Exception
	{
		FileOutputStream Of = new FileOutputStream(srcfile);
		int i;
		int j;
		
		int rc = s.getLastRowNum() - s.getFirstRowNum();
	  	System.out.println("rc "+rc);
	  	  	
	  	for(i=5;i<rc+1;i++)
	  	{
	  		Row row = s.getRow(i);
	  		for(j=0;j<row.getLastCellNum();)
	  		{
	  			String flag = s.getRow(i).getCell(j).getStringCellValue();
	  			
	  			if(flag.equals("Y"))
	  			{
		  			w.findElement(By.partialLinkText("ADD")).click();
		  			
		  			j=j+2;
		  			scrip = s.getRow(i).getCell(j).getStringCellValue();
		  			w.findElement(By.id("search_list")).sendKeys(scrip);
		  			
		  			j=j+1;
		  			String ME = s.getRow(i).getCell(j).getStringCellValue();
		  			w.findElement(By.id("me")).sendKeys(ME);
		  				  				  			
		  			j=j+1;
		  			String IT = s.getRow(i).getCell(j).getStringCellValue();
		  			w.findElement(By.id("inst_type")).sendKeys(IT);
		  				  			
		  			w.findElement(By.partialLinkText("SUBMIT")).click();	  								  						  								
	
		  			//to check if scrip exist
		  			if(w.getPageSource().contains(scrip.toUpperCase()))
		        	{	  				
		  				getdate();
		  				
		        	}
		  			else
		  			{
		  				w.findElement(By.partialLinkText("CANCEL")).click();
			  			log.log(LogStatus.INFO, "No matching stocks found for the Scrip " +scrip.toUpperCase());
			  			
			  			j=j+1;
			  			rw = s.getRow(i);
			  			cell = rw.createCell(j);
			  			cell.setCellValue("null");
		  				break;
		  			}
		  			
		  			if(ME.equals("NSE") || ME.equals("BSE"))
		  			{
			  			String space = " ";
			  			combine = scrip.toUpperCase().concat(space).concat(IT).concat(space).concat(ME);
			  			System.out.println("combine "+combine);
		  			}
		  			else
		  			{
		  				String space = " ";
			  			combine = scrip.toUpperCase().concat(space).concat(IT).concat(space).concat(ME).concat(space).concat(getexactexpdate);
			  			System.out.println("combine "+combine);
		  			}
		  			
		  			j=j+1;
		  			if(ME.equals("NSE-Derivatives"))
		  			{
		  				ME="NSE";
		  				String space = " ";
			  			combine = scrip.toUpperCase().concat(space).concat(IT).concat(space).concat(ME).concat(space).concat(getexactexpdate);

		  			}
		  			if(ME.equals("BSE-Derivatives"))
		  			{
		  				ME="BSE";
		  				String space = " ";
			  			combine = scrip.toUpperCase().concat(space).concat(IT).concat(space).concat(ME).concat(space).concat(getexactexpdate);

		  			}
		  					  			
		  			rw = s.getRow(i);
		  			cell = rw.createCell(j);
		  			cell.setCellValue(combine);
		  					  			
		  			j=j+1;		  			  			
		  			rw = s.getRow(i);
		  			cell = rw.createCell(j);
		  			cell.setCellValue(b);		  			  			
		  			
		  			log.log(LogStatus.INFO, "Scrip "+scrip.toUpperCase()+ " got added.");
		  			break;
		  		}
	  			else
	  			{
	  				break;
	  			}
	  		}
	  	}
	  	wb.write(Of);
	  	Of.close();
	}
	
	
	@Test(priority=458)
	public void nouse()
	  {
		  //report1.endTest(log);
		  //report1.flush();
		  //w.get("D:\\Saleema\\Manage.html");
	  }
}
