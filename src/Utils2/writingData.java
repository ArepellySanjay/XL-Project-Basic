package Utils2;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class writingData {

	public static void main(String[] args) throws IOException 
	{
		
		
		FileInputStream fi = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		
		Workbook wb = new XSSFWorkbook(fi);
	       Sheet ws =	wb.getSheet("empData");
		   Row row = ws.getRow(0);
		  Cell cell = row.createCell(6);	
		 cell.setCellValue("sanju");
		 
		 FileOutputStream fo = new FileOutputStream("E:\\Qedge\\XL.operation.xlsx");
		  wb.write(fo);
		  wb.close();
		  
		
		

	}

}
