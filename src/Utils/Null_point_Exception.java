package Utils;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class  Null_point_Exception{

	public static void main(String[] args) throws IOException {
	
		
		/*FileInputStream san = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		
		Workbook ws = new XSSFWorkbook(san);
		
	    Sheet ss = ws.getSheet("EmpData");
		 
	      Row row =	ss.getRow(1);	      
	     Cell cell = row.getCell(0);
	     
	     	     
	    String data;
	    try {
			
	    	data = cell.getStringCellValue();
	    	System.out.println(data);
		} catch (Exception e) 
	    {
			System.out.println("no data found");
			
		}   
	      
          ws.close();*/
		
		
		/*FileInputStream fi = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		
		Workbook wb = new XSSFWorkbook(fi);
		
	     Sheet ws =	wb.getSheet("EmpData");
		
	       Row row = ws.getRow(1);
	       
	      Cell cell = row.getCell(0);
	      
	      String empno;
	      try {
			
	    	  
	    	  empno  = cell.getStringCellValue();  
	    	  System.out.println(empno);
	    	  
		} catch (Exception e) 
	    {
			
			
			System.out.println("no data found");
			
		}
	            
	        wb.close();*/
	      
		
		
		FileInputStream fi = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		
		Workbook wb = new XSSFWorkbook(fi);
		
	Sheet ws =	wb.getSheet("EmpData");
		
        Row row =	 ws.getRow(1);
		
	Cell c1 =	row.getCell(0);
		
	String data;
	
	try {
		
		data =	c1.getStringCellValue();
		
		System.out.println(data);
	} catch (Exception e) 
	{
		
		System.out.println("no data found");
	}
	
		wb.close();
		
		
		
		
	}

}
