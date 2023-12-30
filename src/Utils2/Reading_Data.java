package Utils2;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Reading_Data {

	public static void main(String[] args) throws IOException 
	{
		
		FileInputStream fi = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		Workbook wb = new XSSFWorkbook(fi);
       	Sheet ws =	wb.getSheet("EmpData");
		     Row row = ws.getRow(0);
		     Cell cell =  row.getCell(1);
		     cell.getStringCellValue();
		     System.out.println(cell);
		     
	}

}
