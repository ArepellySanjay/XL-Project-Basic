package Utils;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read_all_rows_of_data_present_in_XLsheet {

	public static void main(String[] args) throws IOException {


		/*FileInputStream san = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		Workbook wb = new XSSFWorkbook(san);

	     Sheet  ws = wb.getSheet("EmpData");
	     int rowcount =  ws.getLastRowNum();

		for(int i=1;i<=rowcount;i++)
		{

		Row row =	ws.getRow(i);

		Cell c1 = row.getCell(0);
		Cell c2 = row.getCell(1);
		Cell c3 = row.getCell(2);
	   Cell c4 =  row.getCell(3);

	      String empno = c1.getStringCellValue();
	      String empname  =  c2.getStringCellValue();
	      double salary =   c3.getNumericCellValue();   
		  boolean status = c4.getBooleanCellValue();


		  System.out.println(empno+" "+empname+" "+salary+" "+status);



		}



		 wb.close();*/

		//--------------------------------------------------------------------------------------------		

		/*FileInputStream fi  = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");

		Workbook wb = new XSSFWorkbook(fi);

		Sheet ws = wb.getSheet("EmpData");
		int rowcount = ws.getLastRowNum();
		Row row1 = ws.getRow(0);
		Cell c5 = row1.getCell(0);
		Cell c6 = row1.getCell(1);
		Cell c7 = row1.getCell(2);
		Cell c8 = row1.getCell(3);
		String empno = c5.getStringCellValue();
		String empname = c6.getStringCellValue();
		String salary = c7.getStringCellValue();
		String status = c8.getStringCellValue();
		System.out.println(empno+" "+empname+" "+salary+" "+status);
		for(int i=1;i<=rowcount;i++)
		{

			Row row = ws.getRow(i);

			Cell c1 =	row.getCell(0);
			Cell c2 =	row.getCell(1);
			Cell c3 = row.getCell(2);		
			Cell c4 = row.getCell(3);		

			String empno1 = c1.getStringCellValue();
			String empname1 = c2.getStringCellValue();    
			double salary1 =  c3.getNumericCellValue();
			boolean status1 = c4.getBooleanCellValue();

			System.out.println(empno1+" "+empname1+" "+salary1+" "+status1);


		}





		wb.close();*/


  FileInputStream fi =new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
  
  Workbook wb = new XSSFWorkbook(fi);
  
 Sheet ws=  wb.getSheet("EmpData");

   int sheet_rowcount =  ws.getLastRowNum();
       
    for(int i=1;i<=sheet_rowcount;i++)
    {
       Row row = ws.getRow(i);    	
    	
    	Cell c1 = row.getCell(0);
    	Cell c2 = row.getCell(1);
    	 Cell c3 = row.getCell(2);
    	  Cell c4 = row.getCell(3);
    	 Cell  c5 = row.getCell(4);
    	  
    	 String empno = c1.getStringCellValue();
      String empname =	c2.getStringCellValue();
     double salary =	c3.getNumericCellValue();
    boolean status =	c4.getBooleanCellValue();
     String result = c5.getStringCellValue();
      
    
    
       System.out.println(empno+" "+empname+" "+salary+" "+status+" "+result );  
    }
  
  
  wb.close();
  
  
	}

}
