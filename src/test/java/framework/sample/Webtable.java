package framework.sample;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import javax.swing.text.html.parser.TagElement;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Webtable {
	
	public static void main(String[] args) throws IOException, InterruptedException {
		WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();
		driver.get("https://www.icc-cricket.com/rankings/mens/team-rankings/odi");
		Thread.sleep(3000);
		driver.manage().window().maximize();
		WebElement table = driver.findElement(By.xpath("//table[@class='table']"));
		List<WebElement> tableRow = table.findElements(By.tagName("tr"));
		File f = new File("C:\\Users\\Aravindh\\Downloads\\Framework\\IPLTable.xlsx");
		FileInputStream stream1 = new FileInputStream(f);
		Workbook w = new XSSFWorkbook();
		Sheet sheet = w.createSheet("IPL");
		for (int i = 0; i < tableRow.size(); i++) {
			WebElement webElement = tableRow.get(i);
			List<WebElement> value = webElement.findElements(By.tagName("td"));
		for (int j = 0; j < value.size(); j++) {
			WebElement data = value.get(j);
			Row row = null;
			if (j==0) {
				row = sheet.createRow(i);
			} else {
				row = sheet.getRow(i);
			}
			Cell cell = row.createCell(j);
			String text = data.getText();
			cell.setCellValue(text);
		}}
		FileOutputStream stream = new FileOutputStream(f);
		w.write(stream);
	}
}
