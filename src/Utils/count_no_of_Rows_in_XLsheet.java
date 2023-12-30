package Utils;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class count_no_of_Rows_in_XLsheet 
{

	public static void main(String[] args) throws IOException
	{
	
/* FileInputStream fi = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		 
	Workbook wb = new XSSFWorkbook(fi);
	
 Sheet ws1 = wb.getSheet("LoginData");
	
  Sheet ws2 = wb.getSheet("EmpData");
  
   
   int sheet1_rowcount =  ws1.getLastRowNum();
   
 int sheet2_rowcount = ws2.getLastRowNum();
   
  System.out.println("Sheet1 row count:"+sheet1_rowcount);
  System.out.println("Sheet2 row count:"+sheet2_rowcount);
  
  wb.close();*/
	
//-------------------------------------------------------------------------------------
 /* 
  FileInputStream fi = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
  
  Workbook wb = new XSSFWorkbook(fi);
  
     Sheet ws1 = wb.getSheet("LoginData");
  
      Sheet ws2 = wb.getSheet("EmpData");
  
  
     int sheet1_rowcount =  ws1.getLastRowNum();
  
      int sheet2_rowcount=  ws2.getLastRowNum();
      
      System.out.println(sheet1_rowcount);
      
      System.out.println(sheet2_rowcount);
      
      wb.close();*/
  //-----------------------------------------------------------------------------------------
  
  
 /* FileInputStream fi = new FileInputStream();
  
  Workbook wb = new XSSFWorkbook(fi);
   
  Sheet ws1 = wb.getSheet("LoginData");
  
  Sheet ws2 =  wb.getSheet("EmpData");
  
  int sheet1_rowcount = ws1.getLastRowNum();
  
 int sheet2_rowcount = ws2.getLastRowNum();
  
  System.out.println(sheet1_rowcount);
  System.out.println(sheet2_rowcount);
  
  wb.close();*/
  
  
 //------------------------------------------------------------------------------------------ 
  
  
		/*FileInputStream fi = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		
		Workbook wb = new XSSFWorkbook(fi);
		
    Sheet ws =	wb.getSheet("EmpData");
		
	int sheet1_rowcount =	ws.getLastRowNum();
		
		System.out.println(sheet1_rowcount);*/
		
		
		FileInputStream fi = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		
		Workbook wb = new XSSFWorkbook(fi);
		
	Sheet ws = 	wb.getSheet("LoginData");
   Sheet  ws2 =		wb.getSheet("EmpData");
         
      int sheet1_rowcount = ws.getLastRowNum();
        
      int sheet2_rowcount = ws2.getLastRowNum();
      
       System.out.println(sheet1_rowcount);
       System.out.println(sheet2_rowcount);
       
	}
	

}
