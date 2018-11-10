package Libraries;

import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.WebElement;

import Main.beforesuitedealer;

public class variables extends beforesuitedealer
{
	
	public HSSFSheet s;
	public String scrip,getscrip,clientcode,getprice,enterscrip,action,getamt;
	public int qty,modqty;
	public float perc,multiple;
	public Double total,getamt1;
	public Row rw;
	public Cell cell;
	public Double roundoff,reqdroundoff,roundoff1,MRM;
	public double matchround,modmatchround;
	List < WebElement > rc1;
	public String totalmargin,marginutilised,marginavailable,totmar,bprice,sprice;
	
	//for utilisation sheet data
	public static int wputil=13; ///write place utilisation
	public static int wpreqmar=wputil+1;
	public static int wpmarutil=wpreqmar+1;
	public static int wptotmar=wpmarutil+1;
	public static int wptotmarutil=wptotmar+1;
	public static int wpmaravail=wptotmarutil+1;
	public static int wpactmaravail=wpmaravail+1;
	public static int wmutil=wpactmaravail+1; ///write modify utilisation
	public static int wmreqmar=wmutil+1;
	public static int wmtotmar=wmreqmar+1;
	public static int wmtotmarutil=wmtotmar+1;
	public static int wmmaravail=wmtotmarutil+1;
	public static int wmactmaravail=wmmaravail+1;
	public static int wcutil=wmactmaravail+1; ///write cancel utilisation
	public static int wctotmar=wcutil+1;
	public static int wctotmarutil=wctotmar+1;
	public static int wcmaravail=wctotmarutil+1;
	public static int wcactmaravail=wcmaravail+1;
	public static int wclastutil=wcactmaravail+1;
	public static int wstatus=wclastutil+1;
	
				
	///////for tc1, tc2 and tc3
	public String srcfile = "D:\\Saleema\\KotakSecurities\\testdata.xls";
	
	public String stckno;
	public String popup1;
	
	public float fixmultiple;
	public int quantity;
	public Double roundoff2;
	public Double roundoff3;
	public Double roundoff4;
	public Double roundoffpr;
	public Double RM;		
	public String capture;
	public String scripname;
	public Double cap;
	public double newroundoff;
	public Double price;
	public Double newamt;	
	public Double tot;	
	public String stc;
	public String scripchk;		
	public String sd;
	public String scrip1;
	public String pr1;	
	public String Status1;
	public String status;
	
	public int i,j,k,rc;	
	
	public String watchlistname;
	public String chk;
	public WebElement ele;
	public String combine;		
	public String forreportscripname;
	public double capbuyamt;
	public String total1;
	
	///for tc4
	public String dte;
	public String cmprdate;
	public String cmprmonth;
	public String cmpryr;
	public String getexactexpdate;
	public String b;
	public HSSFWorkbook wb;
	
	///for tc5
	public double HLprice;	
	public double calHLprice; //action high low price
	public double caldiff;
	public double calLtp;
	public double calLtp1;
	public double calLtp2;
	public double adddiffLtp;
	public double reqmargin;
	public double addcat;
	public double newreqmargin;
	public double rm1;
	public double mrm1;
	public double modreqmargin;

	////m2m
	public int m2mstatus = 17;
	
	
}


