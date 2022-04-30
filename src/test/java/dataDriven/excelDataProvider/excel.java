package dataDriven.excelDataProvider;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

@Test
public class excel {
	public void getExcel() throws IOException {
		// Object[][] data = { { "hello", "text", "1" }, { "byhe", "message", "143" }, {
		// "solo", "call", "3" } };
		// return data;
		// every row in excel should be sent into 1 array
		FileInputStream fis = new FileInputStream("C:\\Users\\lawencon\\eclipse-workspace1\\dataDriven.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheetAt(0);
		int rowCount = sheet.getPhysicalNumberOfRows();
		XSSFRow row = sheet.getRow(0);
		int colomCount = row.getLastCellNum();
		Object data[][] = new Object[rowCount - 1][colomCount - 1];
		for (int i = 0; i < rowCount - 1; i++) {
			System.out.println("Outer loop started");
			row = sheet.getRow(i + 1);
			for (int j = 0; j < colomCount; j++) {
				// data[][] row.getCell(j);
				System.out.println(row.getCell(j));
			}
			System.out.println("Outer loop ended");
		}
	}
}
