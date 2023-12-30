package Utils;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class count_no_of_coloums_in_a_XLsheet_row {

	public static void main(String[] args) throws IOException 
	{
		
		/*FileInputStream ram = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		
		Workbook wb = new XSSFWorkbook(ram);
		
       Sheet ws = wb.getSheet("LoginData");
		 
     // Row row1 = ws.getRow(0);
	  Row row2 = ws.getRow(2);
      
     //short row1_colcount =  row1.getLastCellNum();
     short row2_colcount = row2.getLastCellNum();
		
   //   System.out.println(row1_colcount); 
      System.out.println(row2_colcount);
      
      wb.close();
  
	}*/

	//--------------------------------------------------------------------------------------------
		
	/*	FileInputStream san = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		
		Workbook ram = new XSSFWorkbook(san);
		
	Sheet ws = ram.getSheet("LoginData");
	  
    Row row1 = ws.getRow(0);
   Row row2 = ws.getRow(1);
	
  int row1_colcount = row1.getLastCellNum();
   
 int row2_colcount =	row2.getLastCellNum();
	
	System.out.println(row1_colcount);
	System.out.println(row2_colcount);

	ram.close();*/
	
//_____________________________________________________________________________________________	
	
	/*  FileInputStream ss = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
	  
	  Workbook san = new XSSFWorkbook(ss);
	  
	 Sheet ws = san.getSheet("LoginData");
	Row row1 = ws.getRow(0);
	Row row2 =	ws.getRow(1);
	
short row1_colcount =	row1.getLastCellNum();
short row2_colcount =   row2.getLastCellNum();
		
		System.out.println(row1_colcount);
		System.out.println(row2_colcount);
    san.close();	*/	
		
//----------------------------------------------------------------------------------------------
   /* FileInputStream san = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
    
    Workbook ss = new XSSFWorkbook(san);
    
   Sheet ws = ss.getSheet("EmpData");
	Row row1 = 	ws.getRow(0);
	Row row2 =	ws.getRow(1);
		
	short row1_colcount =	row1.getLastCellNum();
    short row2_colcount = row2.getLastCellNum();
    
    
    System.out.println("Row coloum count"+ row1_colcount);
    System.out.println(row2_colcount);
	ss.close();	*/
		
	//--------------------------------------------------------------------------------	
	
		/*FileInputStream fi = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		Workbook wb = new XSSFWorkbook(fi);
		
	       Sheet ws =	wb.getSheet("EmpData");
		
		   Row row1 =  ws.getRow(0);
		   Row row2 = ws.getRow(1);
		   
		  short row1_colcount =  row1.getLastCellNum();
		   short row2_colcount =  row2.getLastCellNum();
		
		System.out.println(row1_colcount);
		System.out.println(row2_colcount);*/
		
		
		/*FileInputStream fi = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		Workbook wb = new XSSFWorkbook(fi);
		
	Sheet ws =	wb.getSheet("EmpData");
	
         Row row1 =	ws.getRow(0);
      Row row2 =   ws.getRow(1);
		
	short row1_colcount =	row1.getLastCellNum();
  short row2_colcount =	row2.getLastCellNum();
		
		System.out.println(row1_colcount);
		System.out.println(row2_colcount);
		
		wb.close();
		*/
		
		
		/*FileInputStream  fi =  new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		
		Workbook wb = new XSSFWorkbook(fi);
		
	Sheet ws =	wb.getSheet("EmpData");
	 Row row1 = ws.getRow(0);
	Row row2 = ws.getRow(1);
	  
 short row1_colcount =	row1.getLastCellNum();
   short row2_colcount = row2.getLastCellNum();
	 
   System.out.println(row1_colcount);
   System.out.println(row2_colcount);
   
   
   
   
	wb.close();*/
		
//--------------------------------------------------------------------------------------
		
		/*FileInputStream fi = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		
		Workbook wb = new XSSFWorkbook(fi);
		
	Sheet ws =	wb.getSheet("EmpData");
		
	Row  row1 =	ws.getRow(0);
	 Row row2 =  ws.getRow(1);
	 
	short row1_colcount =	row1.getLastCellNum();
	 short row2_colcount =	row2.getLastCellNum();
		
		System.out.println(row1_colcount);
		System.out.println(row2_colcount);
		
		wb.close();*/
		
		
		FileInputStream fi = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		
		Workbook wb = new XSSFWorkbook(fi);
		
	Sheet ws =	wb.getSheet("EmpData");
		
	  Row row1 = ws.getRow(0);
	 Row row2 = ws.getRow(1);
	 Row row3 = ws.getRow(2);
	 Row row4 = ws.getRow(3);
	 
	 
	short row1_colcount= row1.getLastCellNum();
	 short row2_colcount =   row2.getLastCellNum();
	  short row3_colcount = row3.getLastCellNum();
	  short row4_colcount =  row4.getLastCellNum();
	 
	 System.out.println(row1_colcount);
	 System.out.println(row2_colcount);
	 System.out.println(row3_colcount);
	 System.out.println(row4_colcount);
	 
	}	
	
}
