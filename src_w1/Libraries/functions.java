package Libraries;

import java.awt.Robot;
import java.awt.event.KeyEvent;
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

import Main.beforesuitedealer;


public class functions extends beforesuitedealer
{
	Robot robot;
	
	variables vr = new variables();

	public void openWL() throws Exception
	{				  
		FileInputStream f = new FileInputStream(srcfile);
		 wb = new HSSFWorkbook(f);		
		 vr.s = wb.getSheet("Utilisation");
		 
		  robot = new Robot();	

		///open watchlist
		robot.keyPress(KeyEvent.VK_ALT);
		Thread.sleep(500);
		robot.keyPress(KeyEvent.VK_F);
		Thread.sleep(500);
		robot.keyPress(KeyEvent.VK_O);			  		  	
		Thread.sleep(500);
		robot.keyRelease(KeyEvent.VK_ALT);
		  
		///select watchlist
		String WL = vr.s.getRow(1).getCell(5).getStringCellValue();
		w.findElement(By.id("1013")).click();
		WebElement openwltable = w.findElement(By.name("WatchList")).findElement(By.id("1013"));
		openwltable.findElement(By.xpath("//*[contains(@Name, '"+WL+"')]")).click();
		w.findElement(By.id("1")).click();
		log.log(LogStatus.INFO, "Watchlist opened");
		
		//count no of rows
		WebElement wltable1 = w.findElement(By.id("65280")).findElement(By.id("1000"));
		WebElement colcount = wltable1.findElement(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));		
		vr.rc1 = wltable1.findElements(By.xpath("./*[contains(@ControlType, 'ControlType.Custom')]"));
		System.out.println("row count "+ vr.rc1.size());
		/*
		for(int k=1;k<=rc1.size();k++)
		  {
			  vr.getscrip = wltable1.findElement(By.name("Row " + (k) + ", Column 1")).getText();
			  System.out.println("get scrip " +vr.getscrip);
			  if(vr.scrip.equals(vr.getscrip))
			  {
				  wltable1.findElement(By.name("Row " + (k)+ ", Column 1")).click();
				  System.out.println("scrip got ");
				  vr.getprice = wltable1.findElement(By.name("Row " + (k)+ ", Column 9")).getText();
				  System.out.println("get price "+vr.getprice);
				  
				  //close watchlist
				  close();		
				  
				  //robot.keyPress(KeyEvent.VK_TAB);
				  //robot.keyPress(KeyEvent.VK_F1);
				  
				  break;
			  }
			  else
			  {
				  System.out.println("No scrip ");
			  }				  
		  }	
		  */
	}
	
	public void close() throws Exception
	{
		  Thread.sleep(500);
		  robot.keyPress(KeyEvent.VK_ALT);
		  Thread.sleep(500);
		  robot.keyPress(KeyEvent.VK_F);
		  Thread.sleep(500);
		  robot.keyPress(KeyEvent.VK_C);
		  Thread.sleep(500);
		  robot.keyRelease(KeyEvent.VK_ALT);
		  Thread.sleep(500);
	}
	
}