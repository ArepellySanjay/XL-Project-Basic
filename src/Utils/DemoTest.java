package Utils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DemoTest 
{

	public static void main(String[] args) throws IOException 
	{

		/*FileInputStream fi = new FileInputStream("E:/Qedge/001_Testing/practical/gmail_quality_analysis_ver001.xlsx");
		//FileInputStream fi = new FileInputStream("E:\\Qedge\\001_Testing\\practical\\gmail_quality_analysis_ver001.xlsx");


		Workbook wb = new XSSFWorkbook(fi);
		wb.createSheet("Demosheet123");


		FileOutputStream fo = new FileOutputStream("E:/Qedge/001_Testing/practical/gmail_quality_analysis_ver001.xlsx");

		wb.write(fo);
		wb.close();*/
//------------------------------------------------------------------------------------------
		
		
		FileInputStream fi = new FileInputStream("E:\\Qedge\\XL.operation.xlsx");
		Workbook ws = new XSSFWorkbook(fi);
		
		ws.createSheet("sanju");
		
		FileOutputStream fo = new FileOutputStream("E:\\Qedge\\XL.operation.xlsx");
		ws.write(fo);
		ws.close();
		
		
		
		
		
		
		
		
		
		
		
		
	}

}
