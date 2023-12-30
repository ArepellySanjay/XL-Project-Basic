package Utils2;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class fill_red_colour {

	public static void main(String[] args) throws IOException {
		
		
		FileInputStream fi = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		Workbook wb = new XSSFWorkbook(fi);
		
       	Sheet ws =	wb.getSheet("EmpData");
	    Row row  =	ws.getRow(0);
		   Cell cell =  row.createCell(5);
		 cell.setCellValue("sanju1");
		 
		CellStyle failstyle =    wb.createCellStyle();
		 failstyle.setFillForegroundColor(IndexedColors.RED.getIndex());
		 failstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		  	cell.setCellStyle(failstyle); 
		 
		  	FileOutputStream fo = new FileOutputStream("E:\\Qedge\\XL.operation.xlsx");
		     wb.write(fo);
		     wb.close();
		
	}

}
