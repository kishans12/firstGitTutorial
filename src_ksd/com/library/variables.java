package com.library;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.WebElement;

import Main.BSOrderXL;

public class variables extends BSOrderXL
{
	//for utilisation sheet data
	public static int wputil=16; ///write place utilisation
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
	
	//for compare utilisation sheet compareutil
	public static int cwputil=17; ///compareutil write place utilisation
	public static int cwpactutil=cwputil+1;
	public static int cwpmaxutil=cwpactutil+1;
	public static int cwmutil=cwpmaxutil+1;
	public static int cwmactutil=cwmutil+1;
	public static int cwmmaxutil=cwmactutil+1;
	public static int cwcutil=cwmmaxutil+1;
	public static int cwcactutil=cwcutil+1;
	public static int cwcmaxutil=cwcactutil+1;
	public static int cwclastutil=cwcmaxutil+1;
	public static int cwstatus=cwclastutil+1;
	
	//for sell from existing util
	public static int sfestatus=15;;////sell from existing
	
	//for market depth
	public static int mdstatus=12;;////marketdepth
			
	//for tslo utilisation sheet data
		public static int Twputil=17; ///write place utilisation
		public static int Twpreqmar=Twputil+1;
		public static int Twpmarutil=Twpreqmar+1;
		public static int Twptotmar=Twpmarutil+1;
		public static int Twptotmarutil=Twptotmar+1;
		public static int Twpmaravail=Twptotmarutil+1;
		public static int Twpactmaravail=Twpmaravail+1;
		public static int Twmutil=Twpactmaravail+1; ///write modify utilisation
		public static int Twmreqmar=Twmutil+1;
		public static int Twmtotmar=Twmreqmar+1;
		public static int Twmtotmarutil=Twmtotmar+1;
		public static int Twmmaravail=Twmtotmarutil+1;
		public static int Twmactmaravail=Twmmaravail+1;
		public static int Twcutil=Twmactmaravail+1; ///write cancel utilisation
		public static int Twctotmar=Twcutil+1;
		public static int Twctotmarutil=Twctotmar+1;
		public static int Twcmaravail=Twctotmarutil+1;
		public static int Twcactmaravail=Twcmaravail+1;
		public static int Twclastutil=Twcactmaravail+1;
		public static int Twstatus=Twclastutil+1;
		
	///////for tc1, tc2 and tc3
	public String srcfile = "D:\\testfile\\testdata.xls";
	
	public String stckno;
	public String popup1;
	public Row rw;
	public Cell cell;
	public float multiple;
	public float fixmultiple;
	public int quantity;
	public int modqty;
	public Double roundoff1;
	public Double roundoff2;
	public Double roundoff3;
	public Double roundoff4;
	public Double roundoffpr;
	public Double matchround;
	public Double RM;		
	public String capture;
	public String scripname;
	public Double cap;
	public Double roundoff;
	public Double newroundoff;
	public Double price;
	public Double newamt;	
	public Double tot;	
	public Double total;
	public String stc;
	public String scripchk;		
	public String sd;
	public String scrip1;
	public String action;
	public String pr1;	
	public String Status1;
	public String status;
	
	public int i,j,k,rc;	
	
	public HSSFSheet s;
	public String watchlistname;
	public String chk;
	public WebElement ele;
	public String scrip;
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

	////utilisation supermultiple
	public static int category = 21;
	public static int diffcat = category+1;
	public static int smutil=diffcat+1; //wputil
	public static int smreqmar=smutil+1; ///wpreqmar
	public static int smmarutil=smreqmar+1;  ///wpmarutil
	public static int smmutil= smmarutil+1; ///wmutil
	public static int smmreqmar=smmutil+1; ///wmreqmar
	public static int smcutil=smmreqmar+1;  ////wcutil
	public static int smclutil=smcutil+1;  ///wclastutil
	public static int smstatus=smclutil+1;
	
}


