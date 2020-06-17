package Junit;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;


public class DataDrivenClass {
	  
	  WebDriver mydriver;
		String  expMonthlyPayment,myMonthlyPayment;
		String vExecute,vHomeValue,vDownPay,vLoanAmt,vIRate;
		String [][]xTD;
		String xlPath,xlSheet;
		String xlPath_Res;
		int xlRows;
		
		
  
  
  @Before
  public void beforeTest() {
	  System.setProperty("webdriver.chrome.driver","D:\\QA\\Selenium Automation\\1st Class\\chromedriver_win32\\chromedriver.exe");
		mydriver=new ChromeDriver();
		mydriver.navigate().to("http://www.mortgagecalculator.org/");
		mydriver.manage().window().maximize();
		
		try {     
		xlPath="D:\\QA\\Selenium Automation\\READXL\\TestData.xls";
		xlSheet="TestData";
		xTD=readXL(xlPath,xlSheet);
		xlPath_Res="D:\\QA\\Selenium Automation\\READXL\\TestData_Res.xls";
		xlRows=xTD.length;
		System.out.println("Rows are :"+xlRows);
		}catch(Exception e) {
			System.out.println(e.getMessage());
		}
		}
	
  
  
  
  @Test
  public void mytest() {
	 try {
	  System.out.println("###################### Test is Start ##########################");
	  
	  for(int i=1;i<xlRows;i=i+1) {
		  vExecute=xTD[i][0];
		 
		  if(vExecute.equalsIgnoreCase("Y")) {
			  System.out.println("Running for TDID:"+xTD[i][1]);
			vHomeValue=xTD[i][2];
			vDownPay=xTD[i][3];
			vLoanAmt=xTD[i][4];
			vIRate=xTD[i][5];
			expMonthlyPayment=xTD[i][6];
			
			putData();
			xTD[i][8]=compareResult();
			xTD[i][7]=myMonthlyPayment;
			
	  }else {
		  System.out.println("Skipping for TDID: "+xTD[i][1]);
	  }
	  
	  }
	 
  }catch(Exception e) {
	  System.out.println(e.getMessage());
  }
	System.out.println("###################### Test is Done ##########################") ;
  }  
  
  @After
  public void afterTest() { 
	  try {
	  writeXL(xlPath_Res,"TestResults",xTD);
	  mydriver.close();
	  }catch(Exception e) {
		  System.out.println(e.getMessage());
		  
	  }
  }

  
  
  
  public void putData() { 
      
      mydriver.findElement(By.name("param[homevalue]")).clear();
      mydriver.findElement(By.name("param[homevalue]")).sendKeys(vHomeValue);
      
      mydriver.findElement(By.name("param[downpayment]")).clear();
      mydriver.findElement(By.name("param[downpayment]")).sendKeys(vDownPay);
      
      mydriver.findElement(By.id("loanamt")).clear();
      mydriver.findElement(By.id("loanamt")).sendKeys(vLoanAmt);
      
      mydriver.findElement(By.id("intrstsrate")).clear();
      mydriver.findElement(By.id("intrstsrate")).sendKeys(vIRate);
      
      mydriver.findElement(By.name("cal")).click();
     
	  
  }
  
  public String compareResult() {
	  myMonthlyPayment="None";//Default Value
	  myMonthlyPayment= mydriver.findElement(By.xpath("(//h3)[2]")).getText();        
      
      System.out.println("Actual Monthlypayment is:"+myMonthlyPayment);
      System.out.println("Expected Monthlypayment is:"+expMonthlyPayment);
      
      if(myMonthlyPayment.equals(expMonthlyPayment)) {
      	System.out.println("Test Passed");
      	return "Pass";
      }
      else {
      	System.out.println("Test Failed");
      	return "Fail";
      }
      
  }
  
  //##################### Utility Functions #####################################
  
  //Method of ReadXL
  public static String [][] readXL(String fPath,String fSheet) throws Exception{
	    String [][]xData;
	    int xRows,xCols;
	    DataFormatter dataFormatter=new DataFormatter(); 
	    String cellValue;
        File myxl=new File(fPath);
		FileInputStream myStream=new FileInputStream (myxl);
		HSSFWorkbook myWB = new HSSFWorkbook (myStream);
		HSSFSheet mySheet=myWB.getSheet(fSheet);
		xRows= mySheet.getLastRowNum() + 1;
		xCols=mySheet.getRow(0).getLastCellNum();
		xData=new String[xRows][xCols];
		System.out.println("Rows : "+xRows);
		System.out.println("Cols :"+ xCols);
		
		System.out.println("######################  Test Data Below  ################");
		
		for(int i=0;i<xRows;i=i+1) {
		HSSFRow row=mySheet.getRow(i);
		for(int j=0;j<xCols;j=j+1) {
			cellValue="-";
			cellValue=dataFormatter.formatCellValue(row.getCell(j));
			if(cellValue!=null) {
				xData[i][j]=cellValue;
			}
			System.out.print(cellValue);
			System.out.println("|||||");
		}	System.out.println("");  
  }
  myxl=null;
  return xData;
  
  }
  
  
  
  //Method of WriteXL
  public static void writeXL(String fPath,String fSheet,String[][]xData) throws Exception {
	  
	  File outFile=new File(fPath);
	  HSSFWorkbook wb = new HSSFWorkbook();
	  HSSFSheet osheet=wb.createSheet(fSheet);
	  int xR_TS=xData.length;
	  int xC_TS=xData[0].length;
	  for(int myrow=0;myrow<xR_TS;myrow++) {
		  HSSFRow row=osheet.createRow(myrow);
		  for(int mycol=0;mycol<xC_TS;mycol++) {
			  HSSFCell cell=row.createCell(mycol);
			 //cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			  cell.setCellValue(xData[myrow][mycol]);
			  
		  }
		  FileOutputStream fOut=new  FileOutputStream(outFile);
		  wb.write(fOut);
		  fOut.flush();
		  fOut.close();
	  }
	  wb=null;
	  osheet=null;
	  
	  
  }
  
 
}
  

