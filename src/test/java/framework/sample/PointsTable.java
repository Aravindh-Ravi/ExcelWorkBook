package framework.sample;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PointsTable {
	
	public static void main(String[] args) throws IOException {
		
		File f = new File("C:\\Users\\Aravindh\\Downloads\\Framework\\sample.xlsx");
		
		FileInputStream stream = new FileInputStream(f);
		
		Workbook w = new XSSFWorkbook(stream);
		
		Sheet sheet = w.getSheet("course");
		
		Row row = sheet.getRow(0);
		
		Cell cell = row.getCell(0);
		
		String stringCellValue = cell.getStringCellValue();
		boolean equals = stringCellValue.equals("SeleniumFrameWork");
		
		if (equals) {
			cell.setCellValue("JavaSelenium");
		}
		
		FileOutputStream streamOut = new FileOutputStream(f);
		
		w.write(streamOut);
		
	}
	

}
