package Utils;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class count_no_of_Rows_in_XLSheet2 {

	public static void main(String[] args) throws IOException
	{
		
  FileInputStream sanju = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
   Workbook wb = new XSSFWorkbook(sanju);
   
  Sheet ws1 =  wb.getSheet("LoginData");
  Sheet ws2 =   wb.getSheet("EmpData");		
		
   int Sheet1_rowcount = ws1.getLastRowNum();	
   int Sheet2_rowcount = ws2.getLastRowNum();
		
   System.out.println(Sheet1_rowcount);
   System.out.println(Sheet2_rowcount);
   
   wb.close();
		
		
		
		//count number of rows in xl sheet
		
		
		
		/*FileInputStream sanju = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		Workbook san = new XSSFWorkbook(sanju);
		
	Sheet ws1 =	san.getSheet("LoginData");
    Sheet ws2 = san.getSheet("EmpData");
    
    int Sheet1_rowcount = ws1.getLastRowNum();
    int Sheet2_rowcount = ws2.getLastRowNum();
		 
  
  System.out.println("Sheet1 row count:"+Sheet1_rowcount);
  System.out.println(Sheet2_rowcount);
  
  san.close();*/
		
		
   
	}

}
