package Utils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Writing_data_in_XLsheet {

	public static void main(String[] args) throws IOException 
	{
		
	/*FileInputStream san = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");	
		
		Workbook wb = new XSSFWorkbook(san);
		
      Sheet ws = wb.getSheet("EmpData");
     Row row = ws.getRow(1);
       
    Cell cell = row.createCell(4);
		cell.setCellValue("pass");
		
		FileOutputStream sun = new FileOutputStream("E:\\Qedge\\Result.xlsx");
		wb.write(sun);
         wb.close();*/
         
      //---------------------------------------------------------------------------------   
         
       /* FileInputStream fi = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
        
        Workbook wb  = new XSSFWorkbook(fi);
        
       Sheet ws = wb.getSheet("EmpData");
	Row row = 	ws.getRow(2);
		
	Cell cell =	row.createCell(4);
		cell.setCellValue("pass");
		
		
		FileOutputStream fo =  new FileOutputStream("E:\\Qedge\\sanjuArp.xlsx");
			wb.write(fo);
			
			wb.close();*/
		
			
//-------------------------------------------------------------------------------------
			
			FileInputStream fi = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
			
			Workbook wb = new XSSFWorkbook(fi);
		Sheet ws =	wb.getSheet("EmpData");
			 
           Row row = ws.getRow(1);
          Row row1 = ws.getRow(2);
			 
           Cell cell = row.createCell(4);
          Cell cell1 = row1.createCell(4);
          
          
			cell.setCellValue("pass");
			cell1.setCellValue("pass");
			
			
			FileOutputStream fo = new FileOutputStream("E:\\Qedge\\sanjuArp.xlsx");
			wb.write(fo);
			wb.close();
			
			
	}

}
