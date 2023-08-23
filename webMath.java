package Assignmentprjct;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class webMath {
	
	static String path = "C:\\Users\\Harini T M\\eclipse-workspace\\daily\\excel\\Data.xlsx";
	static XSSFWorkbook wb;

public static void access_data(String path,String Sheet_name) {
	
	try {
		wb = new XSSFWorkbook(path);
	} catch (IOException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}

	
	XSSFSheet sheet = wb.getSheet("Sheet1");
	
	int rows = sheet.getPhysicalNumberOfRows();
	
	int cols = sheet.getRow(0).getPhysicalNumberOfCells();
	
Object data[][] = new Object[rows-1][cols];
	
	for(int i = 1; i<rows ; i++) {
		
		for (int j = 0; j<cols; j++) {
			
 XSSFRichTextString cell_data = sheet.getRow(i).getCell(j).getRichStringCellValue();
			data[i-1][j] = cell_data;
		}
	}

}

	
	public static void main(String[] args) {
		
		access_data(path, "sheet1");
	}
	
	}
		


