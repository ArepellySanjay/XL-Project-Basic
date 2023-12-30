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

public class Fil_XL_sheet_cell_colour_Red_fail {

	public static void main(String[] args) throws IOException {
		
		FileInputStream fi  = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		Workbook wb = new XSSFWorkbook(fi);
		
	        Sheet ws = wb.getSheet("EmpData");
	            Row row1 = ws.getRow(1);
		       Row row2 = ws.getRow(2);
		   
		       
		      Cell cell1 = row1.createCell(4);
		    Cell cell2 =  row2.createCell(4);
		    
		    
		    
		    cell1.setCellValue("pass");
		    cell2.setCellValue("fail");
              
		    
		    CellStyle passstyle = wb.createCellStyle();
		    passstyle.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
		    passstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
             cell1.setCellStyle(passstyle);		    
		    
		    
		    
		    CellStyle failstyle = wb.createCellStyle();
		    failstyle.setFillForegroundColor(IndexedColors.RED.getIndex());
		    failstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		    
	         cell2.setCellStyle(failstyle);
			   
			   
			FileOutputStream fo = new FileOutputStream("E:\\Qedge\\Fill.Green.colour.pass.xlsx");
			wb.write(fo);
			wb.close();
		
		
	}

}
