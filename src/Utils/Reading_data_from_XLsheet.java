package Utils;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Reading_data_from_XLsheet {

	public static void main(String[] args) throws IOException {
	
		
		/*FileInputStream sun = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		
		Workbook san = new XSSFWorkbook(sun);
		
	    Sheet ws =	san.getSheet("EmpData");
	    
        Row row = ws.getRow(1);
        
        
      Cell c1=  row.getCell(0);
     Cell c2 =  row.getCell(1);
     Cell c3 = row.getCell(2);
     Cell c4 = row.getCell(3);
     
    String empid = c1.getStringCellValue();
    String empname = c2.getStringCellValue();
    double salary = c3.getNumericCellValue(); 
    boolean status =  c4.getBooleanCellValue(); 
     
     System.out.println(empid+" "+empname+" "+salary+" "+status);
   
     san.close();
      
     /* System.out.println(c1.getStringCellValue());
      System.out.println(c2.getStringCellValue());
     System.out.println(c3.getNumericCellValue());
     System.out.println(c4.getBooleanCellValue());*/
		
		
     
   /*  FileInputStream fi  = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
     
     Workbook wb =  new XSSFWorkbook(fi);
    Sheet ws = wb.getSheet("EmpData");
    
    Row row = ws.getRow(1);
     
  Cell c1 =  row.getCell(0);
  Cell c2 =   row.getCell(1);     
  Cell c3 =  row.getCell(2);  	
  Cell c4 =    row.getCell(3); 
     
   String empno =  c1.getStringCellValue();
  String empname = c2.getStringCellValue();  
  double salary = c3.getNumericCellValue();
   boolean status = c4.getBooleanCellValue();    
     
     System.out.println(empno+" "+empname+" "+salary+" "+status);
     
     wb.close();*/
     
    /* FileInputStream fi = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
     
     Workbook wb = new XSSFWorkbook(fi);
     
   Sheet ws =  wb.getSheet("EmpData");
   
   Row row =   ws.getRow(1);
   
  Cell c1 = row.getCell(0);
 Cell c2 = row.getCell(1);
  Cell c3 = row.getCell(2);
 Cell c4 = row.getCell(3);
 
String empno =  c1.getStringCellValue();
 String empname = c2.getStringCellValue();
  double salary = c3.getNumericCellValue(); 
    boolean status = c4.getBooleanCellValue();
     
     System.out.println(empno+" "+empname+" "+salary+" "+status); 
     
     wb.close();*/
		
     
     FileInputStream fi = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
     Workbook wb = new XSSFWorkbook(fi);
     
   Sheet ws =  wb.getSheet("EmpData");
    Row row = ws.getRow(1);
    
    Cell c1 = row.getCell(0);
    Cell c2 = row.getCell(1);
      Cell c3 = row.getCell(2);
      Cell c4 = row.getCell(3);
      
     String empno = c1.getStringCellValue();
      String empname = c2.getStringCellValue();
     double salary =  c3.getNumericCellValue();
     boolean status = c4.getBooleanCellValue();
    /* 
     System.out.println(empno+" "+empname+" "+salary+" "+status);*/
     
     /* System.out.println(c1.getStringCellValue());
      System.out.println(c2.getStringCellValue());
      System.out.println(c3.getNumericCellValue());
      System.out.println(c4.getBooleanCellValue());
      */
      
      System.out.println(empno);
      
      
     wb.close();
     
	}

}
