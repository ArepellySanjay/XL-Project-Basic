package Utils;

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

public class Fil_XL_sheet_Cell_colourGreen_pass {

	public static void main(String[] args) throws IOException 
{
		// TODO Auto-generated method stub
		
		/*FileInputStream fi = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		Workbook wb = new XSSFWorkbook(fi);
		
	       Sheet ws =	wb.getSheet("EmpData");
		
	       Row row = ws.getRow(1);
	       
	       Cell cell = row.createCell(4);
	       cell.setCellValue("fail");
		
	       CellStyle failstyle = wb.createCellStyle();
	       
	       failstyle.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
	       failstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	       
	       cell.setCellStyle(failstyle);
	      
	       
	       FileOutputStream sun = new FileOutputStream("E:\\Qedge\\Resultsanji.xlsx");
			wb.write(sun);
			
		
	         wb.close();*/
	       
	       
		//Script to fill XLsheet cell colour with Green
		
		FileInputStream fi  = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		Workbook wb = new XSSFWorkbook(fi);
		
	        Sheet ws = wb.getSheet("EmpData");
		       Row row = ws.getRow(1);
		   
		    Cell cell =  row.createCell(4);
		   cell.setCellValue("pass");
		
		   CellStyle passstyle = wb.createCellStyle();
		   passstyle.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
		   passstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		   
		   cell.setCellStyle(passstyle);
		   
		   
		FileOutputStream fo = new FileOutputStream("E:\\Qedge\\Fill.Red.colour.fail.xlsx");
		wb.write(fo);
		wb.close();

		
		
		
	}

}
