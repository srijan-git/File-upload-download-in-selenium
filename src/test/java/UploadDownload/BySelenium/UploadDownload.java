package UploadDownload.BySelenium;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

public class UploadDownload {
	public static void main(String[] args) throws IOException {

		String fruiteName = "Apple";

		String fileName = "/home/srijan-kumar-khan/Downloads/download.xlsx";

		String updatedValue = "598";

		WebDriver driver = new ChromeDriver();

		driver.get("https://rahulshettyacademy.com/upload-download-test/index.html");

		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(7));

		driver.manage().window().maximize();

		// Download
		driver.findElement(By.cssSelector("#downloadButton")).click();

		// Edit excel
		int columnNumber = getColumnNumber(fileName, "price");

		int rowNumber = getRowNumber(fileName, fruiteName);

		Assert.assertTrue(updateCell(fileName, rowNumber, columnNumber, updatedValue));

		// Upload
		WebElement uploadElement = driver.findElement(By.cssSelector("input[type='file']"));

		uploadElement.sendKeys("/home/srijan-kumar-khan/Downloads/download.xlsx");

		// Wait for success messahe to show up and wait for disappear

		By toastLocator = By.cssSelector(".Toastify__toast-body div:nth-child(2)");
		WebDriverWait driverWait = new WebDriverWait(driver, Duration.ofSeconds(10));

		driverWait.until(ExpectedConditions.visibilityOfElementLocated(toastLocator));

		String actualToastString = driver.findElement(toastLocator).getText();

		Assert.assertEquals("Updated Excel Data Successfully.", actualToastString);

		driverWait.until(ExpectedConditions.invisibilityOfElementLocated(toastLocator));

		// verify updated excel data showing in the web table
		String priceColumn = driver.findElement(By.xpath("//div[text()='Price']")).getAttribute("data-column-id");

		String actualPriceString = driver.findElement(By.xpath("//div[text()='" + fruiteName
				+ "']/parent::div//parent::div/div[@id='cell-" + priceColumn + "-undefined']")).getText();

		Assert.assertEquals(updatedValue, actualPriceString);
	}

	private static int getRowNumber(String fileName, String text) throws IOException {

		FileInputStream fileInputStream = new FileInputStream(fileName);

		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

		XSSFSheet sheet = workbook.getSheet("Sheet1");

		Iterator<Row> rows = sheet.iterator();

		int k = 1;

		int rowIndex = -1;

		while (rows.hasNext()) {
			Row row = rows.next();
			Iterator<Cell> cell = row.cellIterator();
			while (cell.hasNext()) {
				Cell valueCell = cell.next();
				if (valueCell.getCellType() == CellType.STRING
						&& valueCell.getStringCellValue().equalsIgnoreCase(text)) {
					rowIndex = k;
				}
			}
			k++;
		}

		return rowIndex;
	}

	private static int getColumnNumber(String fileName, String colName) throws IOException {

		FileInputStream fileInputStream = new FileInputStream(fileName);

		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

		XSSFSheet sheet = workbook.getSheet("Sheet1");

		Iterator<Row> rows = sheet.iterator();
		Row firstRow = rows.next();
		Iterator<Cell> cell = firstRow.cellIterator();
		int k = 1;
		int column = 0;
		while (cell.hasNext()) {
			Cell value = cell.next();
			if (value.getStringCellValue().equalsIgnoreCase(colName)) {
				column = k;
			}
			k++;
		}
		return column;
	}

	private static boolean updateCell(String fileName, int rowNumber, int columnNumber, String updatedValue)
			throws IOException {

		FileInputStream fileInputStream = new FileInputStream(fileName);

		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

		XSSFSheet sheet = workbook.getSheet("Sheet1");

		Row rowField = sheet.getRow(rowNumber - 1);

		Cell cellField = rowField.getCell(columnNumber - 1);

		cellField.setCellValue(updatedValue);

		FileOutputStream fileOutputStream = new FileOutputStream(fileName);

		workbook.write(fileOutputStream);

		workbook.close();

		fileInputStream.close();

		return true;
	}

}
