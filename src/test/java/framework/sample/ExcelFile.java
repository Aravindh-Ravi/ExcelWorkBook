package framework.sample;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFile { 
	
	public static void main(String[] args) throws IOException {
		
		// create object for File
				File f = new File("C:\\Users\\Aravindh\\Downloads\\Framework\\sample.xlsx");
				
				// to read file
				FileInputStream stream = new FileInputStream(f);
				
				// to read the excel
				Workbook w = new XSSFWorkbook(stream);
				
				// to read the sheet
				Sheet sheet = w.getSheet("client data");
				
				// to find the number of rows filled with data
				int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
				
				// to iterate the sperate each rows
				for (int i = 0; i < physicalNumberOfRows; i++) {
					
				Row row = sheet.getRow(i);
					
				// to find the number of cell present in each row
				int physicalNumberOfCells = row.getPhysicalNumberOfCells();
				
				// iterate to seperate each cell
				
				for (int j = 0; j < physicalNumberOfCells; j++) {
					
					Cell cell = row.getCell(j);
					
					int cellType = cell.getCellType();
					
					if (cellType==1) {
						
						String stringCellValue = cell.getStringCellValue();
						
						
					}
					
					else { DateUtil.isCellDateFormatted(cell);
						
						Date dateCellValue = cell.getDateCellValue();
					
					System.out.println(dateCellValue);
					
					SimpleDateFormat d = new SimpleDateFormat("MM/dd/yyyy");
					
					String format = d.format(dateCellValue);
					System.out.println(format);
						
					double numericCellValue = cell.getNumericCellValue();
						
						long l =(long)numericCellValue;
						
						System.out.println(l);
					}
					
					
				}
				}
			
				}
		
	}
