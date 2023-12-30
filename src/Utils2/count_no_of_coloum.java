package Utils2;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class count_no_of_coloum {

	public static void main(String[] args) throws IOException 
	{
	
       FileInputStream fi = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		
		Workbook wb = new XSSFWorkbook(fi);
		Sheet ws = wb.getSheet("EmpData");
         Row row = ws.getRow(0);
         
	       int row1 = 	row.getLastCellNum();
	            int rowc =    ws.getLastRowNum();
		System.out.println(row1);
		System.out.println(rowc);
		
		
	}

}
