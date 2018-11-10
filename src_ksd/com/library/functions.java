package com.library;

import java.io.FileInputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
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

import com.relevantcodes.extentreports.LogStatus;

import Main.BSOrderXL;
import Utilisation.ManageWatchlist;

public class functions extends ManageWatchlist
{
	public static int i;
	public static String sd;
	public static Row rw;
	public static Cell cell;
	public static String b;

	public static String forreportscripname;
	public static String symbol;
	public static String getscrip;
	public static String chk;
	public static WebElement ele;
	
	public void getdate()
	{				
		List<WebElement> myList= w.findElements(By.xpath(".//td[input[@type = 'checkbox']]"));
		List<WebElement> myListele = w.findElements(By.xpath("//input[@name='stockcode' and @type = 'checkbox']"));
		System.out.println("Total elements :" +myList.size());
		
        Date currDate = Calendar.getInstance().getTime();
	   	SimpleDateFormat dateFormat = new SimpleDateFormat("ddMMMyy");
		SimpleDateFormat dateFormatformonth = new SimpleDateFormat("MMM");
		SimpleDateFormat dateFormatforyear = new SimpleDateFormat("yy");
		
		String getcurrdate = dateFormat.format(currDate);
		System.out.println("curr date "+getcurrdate);
	    	 
		Calendar calendar = Calendar.getInstance(TimeZone.getDefault());
		int year = calendar.get(Calendar.YEAR);
		int date = calendar.get(Calendar.DAY_OF_MONTH);
		String month = dateFormatformonth.format(currDate);
		String yr = dateFormatforyear.format(currDate);
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
				cmpryr = getonlydate[2];
				
				int result = dte.compareTo(cmprdate);
				System.out.println("result "+result);
				//if(result<=0 || result>0)
					
				////for fno and currency contract dates
			        fnodate = s2.getRow(8).getCell(2).toString();
			        System.out.println("fno date "+fnodate);
					String[]removehypfno = fnodate.split("\\-");
					String[]remfno20 = removehypfno[2].split("0" ,0);
					String concatfno = removehypfno[0].concat(removehypfno[1].toUpperCase().concat(remfno20[1]));
			        
					currdate = s2.getRow(8).getCell(3).toString();
					String[]removehypcurr = currdate.split("\\-");
					String[]remcurr20 = removehypcurr[2].split("0" ,0);
					String concatcurr = removehypcurr[0].concat(removehypcurr[1].toUpperCase().concat(remcurr20[1]));
					
				if(removebrac[1].equals(concatfno) || removebrac[1].equals(concatcurr))
				{
					//if(month.toUpperCase().equals(cmprmonth)) //comment
        			{
						if(yr.equals(cmpryr))
						{
							b = ele.getAttribute("value");
				  			System.out.println("getvalue " +b);
				  			w.findElement(By.partialLinkText("SUBMIT")).click();
				  			break;
	        			}
        			}
					//else //comment
					{
						//ele.click(); //comment
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
	
	
	public void margincalc() throws Exception
	{
		HSSFSheet s = wb.getSheet("Utilisation");
		//i=Normalallmatch.i;
		
		///click on cancel
		w.findElement(By.partialLinkText("CANCEL")).click();
		
		///click on total utilised cashNfno
		String totalmargin = w.findElement(By.xpath("//*[@id=\'limit-node-1\']/td[3]")).getText();
		String t = totalmargin.replaceAll(",","").trim();
		String m = t.replaceAll("","");
		System.out.println(m);
		log.log(LogStatus.INFO, "Total Margin is " +totalmargin);
		
		System.out.println("margincal i ....."+i);
		rw = s.getRow(i);
		cell = rw.createCell(14);
		cell.setCellValue(totalmargin);
		
		///click on margin utilised cashNfno
		String marginutilised = w.findElement(By.xpath("//*[@id=\'limit-node-3\']/td[3]/a")).getText();
		String t1 = marginutilised.replaceAll(",","").trim();
		String m1 = t1.replaceAll("","");
		System.out.println(m1);
		log.log(LogStatus.INFO, "Margin Utilised is " +marginutilised);
		
		System.out.println("margincal i ....."+i);
		rw = s.getRow(i);
		cell = rw.createCell(15);
		cell.setCellValue(marginutilised);
		
		///click on margin available cashNfno
		String marginavailable = w.findElement(By.xpath("//*[@id=\'limit-node-6\']/td[3]/strong")).getText();
		String t2 = marginavailable.replaceAll(",","").trim();
		String m2 = t2.replaceAll("","");
		System.out.println(m2);
		log.log(LogStatus.INFO, "Margin Available is " +marginavailable);
		
		System.out.println("margincal i ....."+i);
		rw = s.getRow(i);
		cell = rw.createCell(16);
		cell.setCellValue(marginavailable);
		
		//verify totalmargin - margin utilsed = margin available
		Double ab = Double.parseDouble(m)-Double.parseDouble(m1);
		BigDecimal bd = new BigDecimal(ab);
		DecimalFormat df = new DecimalFormat("0.00");
		df.setMaximumFractionDigits(2);
		sd = df.format(ab);
		System.out.println(sd);
		
		System.out.println("margincal i ....."+i);
		rw = s.getRow(i);
		cell = rw.createCell(17);
		cell.setCellValue(sd);
		
		///check if above are equal
		if(sd.equals(m2))
		{
			System.out.println("Equal");
			log.log(LogStatus.INFO, "Utilized -- As Required Margin and Margin Utilised are Equal");
			log.log(LogStatus.INFO, "Total Margin & Margin utilized are equal to Margin Available " +sd);
			
			System.out.println("margincal i ....."+i);
			rw = s.getRow(i);
			cell = rw.createCell(18);
			cell.setCellValue("Pass");
		}
		else
		{
			log.log(LogStatus.INFO, "Total Margin & Margin utilized are equal to Margin Available " +sd);
			
			System.out.println("margincal i ....."+i);
			rw = s.getRow(i);
			cell = rw.createCell(18);
			cell.setCellValue("Fail");

		}
		
	}	
	
}
