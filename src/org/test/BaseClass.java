package org.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

import com.fasterxml.jackson.databind.exc.InvalidFormatException;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseClass {
	static WebDriver driver;

	public void getDriver() {
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
	}

	public void loadUrl(String url) {
		driver.get(url);
	}

	public void maximize() {
		driver.manage().window().maximize();
	}

	public WebElement findElementByid(String attributeValue) {
		WebElement element = driver.findElement(By.id(attributeValue));
		return element;

	}

	public WebElement findElementByName(String attributeValue) {
		WebElement element = driver.findElement(By.name(attributeValue));
		return element;
	}

	public WebElement findElementByXpath(String attributeValue) {
		WebElement element = driver.findElement(By.xpath(attributeValue));
		return element;

	}

	public void sendKeys(WebElement element, String Data) {
		element.sendKeys(Data);
	}

	public void click(WebElement element) {
		element.click();

	}

	public String getData(String sheetName, int rownum, int cellnum)
			throws InvalidFormatException, IOException, org.apache.poi.openxml4j.exceptions.InvalidFormatException {
		String data = null;
		File file = new File("C:\\Users\\Siddharthan\\eclipse-workspace\\FrameworkClass\\ExcelSheets\\Book1.xlsx");
		Workbook workbook = new XSSFWorkbook(file);
		Sheet sheet = workbook.getSheet(sheetName);
		Row row = sheet.getRow(rownum);
		Cell cell = row.getCell(cellnum);
		CellType type = cell.getCellType();
		switch (type) {
		case STRING:
			data = cell.getStringCellValue();
			break;
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				data = new SimpleDateFormat("dd-MMM-yy").format(cell.getDateCellValue());
			} else {
				data = BigDecimal.valueOf(cell.getNumericCellValue()).toString();
			}
			break;

		default:
			break;
		}
		return data;
	}

	public void writeData( String sheetName, int rowno, int cellno, String data) throws IOException {
		File file = new File("C:\\Users\\Siddharthan\\eclipse-workspace\\FrameworkClass\\ExcelSheets\\Book1.xlsx");
		FileInputStream stream = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet(sheetName);
		Row row = sheet.getRow(rowno);
		Cell cell = row.createCell(cellno);
		cell.setCellValue(data);
		FileOutputStream out = new FileOutputStream(file);
		workbook.write(out);
	}

	public String attributeValue(WebElement element, String name) {
		String attribute = element.getAttribute(name);
		return attribute;

	}

	public void dropDownByIndex(WebElement element, int num) {
		Select s = new Select(element);
		s.selectByIndex(num);
	}

	public void waittime(long value) {
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
	}
	
}


}
