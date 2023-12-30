package Utils2;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read_all_rows_in_XLsheet {

	public static void main(String[] args) throws IOException
	{
		
		FileInputStream fi = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		
		Workbook wb = new XSSFWorkbook(fi);
		
		Sheet ws = wb.getSheet("EmpData");
		
		 int rowcount =    ws.getLastRowNum();
		
		 for(int i=1;i<=rowcount;i++)
		 {
			 
          	 	Row row =	ws.getRow(i); 
			 
			Cell c1 = row.getCell(0);
          	Cell c2 =   row.getCell(1);
			Cell c3 = row.getCell(2); 
			Cell c4 = row.getCell(3);
			
		String empid =	c1.getStringCellValue();
		String empname = c2.getStringCellValue();
		 double salary =  c3.getNumericCellValue();
		  boolean result =  c4.getBooleanCellValue();
		     
		  
		  System.out.println(empid);
		  System.out.println(empname);
		  System.out.println(salary);
		  System.out.println(result);
		  
		  
		
			
		 }
		 
		 
		
	}

}
