package PlaceOrder;

import java.io.FileInputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;

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

public class AMOOrderUtilisation extends BSOrderXL
{
	public static String capture;
	public static Double cap;
	public static Double matchround;
	public static Double roundoff;
	public static Double newroundoff;
	public static Double price;
	public static Double tot;
	public static Double RM;
	public static String stc;
	public static int multiple;
	public static int quantity;
	public static String sd;
	public static String scrip1;
	public String stckno;
	public String Status1;
	public String status;
			
	public void capturetotal() throws Exception
	{
		FileInputStream f = new FileInputStream("D:\\Saleema\\KotakSecurities\\src\\uticash.xls");
		HSSFWorkbook wb = new HSSFWorkbook(f);
		HSSFSheet s = wb.getSheet("Sheet1");
		
		///click funds >>equity/derivatives
		Actions act1 = new Actions(w);
		act1.moveToElement(w.findElement(By.linkText("Funds"))).click().perform();	
		WebElement Sublink1 = w.findElement(By.linkText("Equity/Derivatives"));
		Sublink1.click();
					
		//capture utilisation total
		w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).click();
		String checkmsg = "No Record Found";
		String scrip1 = s.getRow(7).getCell(1).getStringCellValue();
		if(!w.getPageSource().contains(checkmsg))
		{
			if(w.getPageSource().contains(scrip1))
			{
			System.out.println("in loop scrip1 " +scrip1);
			String capture = w.findElement(By.xpath("//td[text()='"+scrip1+"']/following-sibling::td[6]")).getText();
			System.out.println("capture " +capture);
			cap = Double.parseDouble(capture);
			System.out.println("first/cap " +cap);
			}
			else
			{
				cap = 0.0;
				System.out.println("cap of capture " +cap);
			}
		}
		else
		{
			System.out.println("Out of loop");
		}
				
	}
	
	public void addscrip() throws Exception
	{
		log = report1.startTest("Add Scrip for FNO");
		
		///click on add
		w.findElement(By.partialLinkText("ADD")).click();
		
		FileInputStream f = new FileInputStream("D:\\Saleema\\KotakSecurities\\src\\uticash.xls");
		HSSFWorkbook wb = new HSSFWorkbook(f);
		HSSFSheet s = wb.getSheet("Sheet1");
		
		///search scrip and enter
		String searchscrip = s.getRow(12).getCell(1).getStringCellValue();
		w.findElement(By.id("search_list")).sendKeys(searchscrip);
		System.out.println("data1 " +searchscrip);
		
		//select market exchange
		Select ME = new Select(w.findElement(By.id("me")));
		ME.selectByVisibleText("NSE-Derivatives");
		
		//select instrument type
		Select IT = new Select(w.findElement(By.id("inst_type")));
		IT.selectByVisibleText("FS");
		
		//select expiry date
		Select ED = new Select(w.findElement(By.id("expiry_date")));
		ED.selectByVisibleText("26APR18");
		
		//click on submit
		w.findElement(By.partialLinkText("SUBMIT")).click();
		
		///select scrip to add
		String xcelscrip = s.getRow(16).getCell(1).getStringCellValue();
		System.out.println("data1 " +xcelscrip);
		w.findElement(By.xpath(".//input[@value='"+xcelscrip+"']")).click();
		
		///click submit
		w.findElement(By.partialLinkText("SUBMIT")).click();
				
	}
	
	@Test(priority=0)
	public void PlaceAMOBuyOrder() throws Exception 
	  {
		FileInputStream f = new FileInputStream("D:\\Saleema\\KotakSecurities\\src\\uticash.xls");
		HSSFWorkbook wb = new HSSFWorkbook(f);
		HSSFSheet s = wb.getSheet("Sheet1");
				
		///report creation
		report1 = new ExtentReports("D:\\Saleema\\AMOUtilisationFNO.html");
	  	log = report1.startTest("Place AMO FNO Order");
	  	
	  	////call login function
	  	Login();
	  	
	  	multiple = (int) s.getRow(3).getCell(1).getNumericCellValue();  	
		System.out.println("data " +multiple);
				
		///call capture total function to capture utilisation value
		capturetotal();
		
	  	///click watchlist>>place order
		Actions act = new Actions(w);
		act.moveToElement(w.findElement(By.linkText("Watchlist"))).click().perform();	
		WebElement Sublink = w.findElement(By.linkText("My Watchlists"));
		Sublink.click();
		log.log(LogStatus.INFO, "Click on Watchlist >> My Watchlists");
		
		///call addscrip
		addscrip();
		
		///select scrip
		String xcelscrip = s.getRow(5).getCell(1).getStringCellValue();
		System.out.println("data1 " +xcelscrip);
		w.findElement(By.xpath(".//input[@value='"+xcelscrip+"']")).click();
		
		////click on place order
		w.findElement(By.partialLinkText("PLACE ORDER")).click();
		log.log(LogStatus.INFO, "Click on Place Order");
		
		///get scripname
		String scripname = w.findElement(By.id("stk_name")).getAttribute("value");
		log.log(LogStatus.INFO, "Selected Scrip is " +scripname);
		
		///select buy/sell
		String value = s.getRow(6).getCell(1).getStringCellValue();
		System.out.println("boolean " +value);
		w.findElement(By.xpath(".//input[@value='"+value+"']")).click();
		log.log(LogStatus.INFO, "Selected Value is " +value);
		
		////enter quantity
		int quantity = (int) s.getRow(4).getCell(1).getNumericCellValue();  	
		System.out.println("data1 " +quantity);
		wb.close();
		w.findElement(By.id("stk_qty")).sendKeys(String.valueOf(quantity));
		log.log(LogStatus.INFO, "Entered Quantity is: " +quantity);
				
		///required margin calculation
		String pr = w.findElement(By.id("stk_price")).getAttribute("value");
		price = Double.parseDouble(pr);
		log.log(LogStatus.INFO, "Entered price is " +price);
		RM = ((quantity * price ) / multiple);
		System.out.println("RM " +RM);
		roundoff = (Math.round(RM*100.0)/100.0);
		log.log(LogStatus.INFO, "Required Margin is " +roundoff);
		System.out.println("roundoff " +roundoff);
		
		///check amo
		w.findElement(By.id("chk_amo")).click();
		log.log(LogStatus.INFO, "Select Place AMO checkbox");
		
		///place order
		w.findElement(By.partialLinkText("PLACE ORDER")).click();
		log.log(LogStatus.INFO, "Click on Place Order button");
		
		///alert
		w.switchTo().alert().accept();
		log.log(LogStatus.INFO, "Click OK on alert");
		
		///check for any error
		String popup = "Error";
		if (w.getPageSource().contains(popup))
		{
			String popup1 = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
			System.err.println("Error displayed as " +popup1);
			log.log(LogStatus.ERROR, "" +popup1);
		}
		else
		{	
			////capture order no
			stckno = w.findElement(By.xpath("//*[@id='container']/div/p/span")).getText();
			System.out.println("stckno...... " +stckno);
			log.log(LogStatus.INFO, "Order No is " +stckno);
			
			///click funds >>equity/derivatives
			Actions act1 = new Actions(w);
			act1.moveToElement(w.findElement(By.linkText("Funds"))).click().perform();	
			WebElement Sublink1 = w.findElement(By.linkText("Equity/Derivatives"));
			Sublink1.click();
			log.log(LogStatus.INFO, "Click on Funds >> Equity/Derivatives");
			
			//click margin utilised cash&FNO
			w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).click();
			
			///check scrip and capture amount
			String scrip1 = s.getRow(7).getCell(1).getStringCellValue();
			System.out.println("scrip1 " +scrip1);
			if(w.getPageSource().contains(scrip1))
			{
				///get util amount of specific scrip				
				String total = w.findElement(By.xpath("//td[text()='"+scrip1+"']/following-sibling::td[6]")).getText();
				System.out.println("total " +total);
				tot = Double.parseDouble(total);
				System.out.println("cap amt " +tot);
				Double match = tot - cap;
				System.out.println("match " +match);
				matchround = Math.round(match*100.0)/100.0;
				System.out.println("f amt" +matchround);
				log.log(LogStatus.INFO, "Margin Utilised Cash & FNO " +matchround);
			
			////to check if reqd margin and captured util amt are same
			if(matchround.equals(roundoff))
			{
				///click on cancel
				w.findElement(By.partialLinkText("CANCEL")).click();
				
				///click on total utilised cashNfno
				String totalmargin = w.findElement(By.xpath("//*[@id=\'limit-node-1\']/td[3]")).getText();
				System.out.println(totalmargin);
				String t = totalmargin.replaceAll(",","").trim();
				String m = t.replaceAll("","");
				System.out.println(m);
				log.log(LogStatus.INFO, "Total Margin is " +totalmargin);
				
				///click on margin utilised cashNfno
				String marginutilised = w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).getText();
				System.out.println(marginutilised);
				String t1 = marginutilised.replaceAll(",","").trim();
				String m1 = t1.replaceAll("","");
				System.out.println(m1);
				log.log(LogStatus.INFO, "Margin Utilised is " +marginutilised);
				
				///click on margin available cashNfno
				String marginavailable = w.findElement(By.xpath("//*[@id=\'limit-node-6\']/td[3]/strong")).getText();
				System.out.println(marginavailable);
				String t2 = marginavailable.replaceAll(",","").trim();
				String m2 = t2.replaceAll("","");
				System.out.println(m2);
				log.log(LogStatus.INFO, "Margin Available is " +marginavailable);
				
				//verify totalmargin - margin utilsed = margin available
				Double ab = Double.parseDouble(m)-Double.parseDouble(m1);
				BigDecimal bd = new BigDecimal(ab);
				System.out.println("bd " +bd);
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
				log.log(LogStatus.ERROR, "Total Margin & Margin utilized are not equal to Margin Available " +sd);
			}
		}
			else
			{
				System.out.println("Selected scrip does not match");
			}
	  }
						
	}
	
	@Test(priority=1)
	public void ManageCashAMO() throws Exception 
	  {		
		///report creation
		log = report1.startTest("Change AMO FNO Order");
	  	
		///to check for any error
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
		WebElement Sublink = w.findElement(By.linkText("Order Status"));
		Sublink.click();
		log.log(LogStatus.INFO, "Click on Reports >> Order Status");
		
	  	//	w.findElement(By.partialLinkText("Order Status")).click();	
	  	
		//check stkno
		System.out.println("manage stk no " +stckno);
		w.findElement(By.xpath(".//input[@value='"+stckno+"|N']")).click();
		log.log(LogStatus.INFO, "Click on Order No " +stckno);
		
		String status = w.findElement(By.xpath("//div[@id='container']//td[contains(text(), 'AMO')]")).getText();
		System.out.println("status is " +status);
		String Status1 = "AMO";				
		System.out.println("Order no " +stckno);
						
		///click change order
		w.findElement(By.partialLinkText("CHANGE ORDER")).click();
				
		String check = "Error";
		if(w.getPageSource().contains(check))
		{
			String popup = w.findElement(By.xpath("//*[@id=\'container\']/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td")).getText();
			log.log(LogStatus.ERROR, "" +popup);
			//w.findElement(By.partialLinkText("BACK")).click();
			System.out.println("Order no " +stckno+ " cannot be modified as status is " +status);

		}
		else
		{		
			FileInputStream f = new FileInputStream("D:\\Saleema\\KotakSecurities\\src\\uticash.xls");
			HSSFWorkbook wb = new HSSFWorkbook(f);
			HSSFSheet s = wb.getSheet("Sheet1");
			
			Double newcap = matchround;
			System.out.println("newcap " +newcap);
						
			//update stock no
			w.findElement(By.id("stk_qty")).clear();
			quantity = (int) s.getRow(4).getCell(5).getNumericCellValue();  	
			System.out.println("data1 " +quantity);
			w.findElement(By.id("stk_qty")).sendKeys(String.valueOf(quantity));
			log.log(LogStatus.INFO, "Enter Quantity: " +quantity);
			wb.close();
			
			///modify order
			RM = ((quantity * price ) / multiple);
			System.out.println("RM " +RM);
			Double roundoff1 = (Math.round(RM*100.0)/100.0);
			System.out.println("roundoff " +roundoff1);
			
			///click on chnge order
			w.findElement(By.partialLinkText("CHANGE ORDER")).click();
			log.log(LogStatus.INFO, "Click on Change Order");	
			
			//cnfrm
			w.findElement(By.partialLinkText("Confirm")).click();
			log.log(LogStatus.INFO, "Click Confirm");
			System.out.println("Order no " +stckno+ " is modified successfully");
			log.log(LogStatus.INFO, "Order no " +stckno+ " is modified successfully");
			
			//double modifymargin = cap - newcap;
			//double modifymargin = tot - newroundoff;
			//System.out.println("modified margin1 " +modifymargin);
			//System.out.println("modified margin " +cap);
			log.log(LogStatus.INFO, "Modified required margin is " +roundoff1);
									
		}
		
	  	}
  }
		
	 @Test(priority = 2)
	  public void CancelAMOBuyOrder() throws Exception
	  {
		 ///for report name
		  log = report1.startTest("Cancel AMO FNO Order");
		  
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
			WebElement Sublink = w.findElement(By.linkText("Order Status"));
			Sublink.click();
			log.log(LogStatus.INFO, "Click on Reports >> Order Status");
			
			//check stkno
			w.findElement(By.xpath(".//input[@value='"+stckno+"|N']")).click();
			log.log(LogStatus.INFO, "Click on Order no " +stckno);
			
			///click on cancel
			w.findElement(By.partialLinkText("CANCEL ORDER")).click();
			log.log(LogStatus.INFO, "Click on Cancel");
			
			///check for error
			String trade = "Error";
			if(w.getPageSource().contains(trade))
			{
				String popup = w.findElement(By.xpath("//div[@id='container']//td[contains(text(), 'Can not cancel Order.It is already TRADED')]")).getText();
				System.out.println("status is "+popup);
				log.log(LogStatus.ERROR, "Order no " +stckno+ " cannot be cancelled as " +popup);				
				w.findElement(By.partialLinkText("Back")).click();	
				System.out.println("Order no "+stckno+ " cannot be cancelled as status is " +status);
				log.log(LogStatus.INFO, "Order no " +stckno+ " cannot be modified as status is " +status);
				
			}
			else
			{
				//cnfrm
				w.findElement(By.partialLinkText("Confirm")).click();
				log.log(LogStatus.INFO, "Click on Confirm");
				System.out.println("Order no "+stckno+ " is cancelled successfully");
				log.log(LogStatus.INFO, "Order no " +stckno+ " is cancelled successfully");
				
				Double newamt = cap;
				System.out.println("newamt " +newamt);
				
				///call capture total
				capturetotal();
				Double canamt = newamt - cap;
				System.out.println("can amt " +canamt);
				log.log(LogStatus.INFO, "Cancelled Required Margin is " +canamt);
			}
		  
		  }
	  }
	 
	@Test(priority=45)
	public void nouse()
	  {
		  report1.endTest(log);
		  report1.flush();
		  w.get("D:\\Saleema\\AMOUtilisationFNO.html");
	  }
}
