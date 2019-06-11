package Package_test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;

public class ApachePOITest {
	static List<String> usernames = new ArrayList<String>();
	List<String> password = new ArrayList<String>();
	static FileInputStream file;
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	static XSSFCell cell;
	XSSFRow row;

	@BeforeClass
	public static void setup() throws IOException {
		file = new FileInputStream(
				"C:\\Users\\Admin\\Desktop\\EclipseWorkspace\\Excel\\src\\main\\resources\\DemoSiteDDT.xlsx");
		workbook = new XSSFWorkbook(file);
		sheet = workbook.getSheetAt(0);

		// Reading
		for (int rowNum = 0; rowNum < sheet.getPhysicalNumberOfRows(); rowNum++) {
			for (int colNum = 0; colNum < sheet.getRow(rowNum).getPhysicalNumberOfCells(); colNum++) {
				cell = sheet.getRow(rowNum).getCell(colNum);
				String userCell = cell.getStringCellValue();
				System.out.println(userCell);
				usernames.add(userCell);
				// assertEquals(“Expected”, userCell);
			}
		}

		file.close();
	}

	@Test
	public void test1() {

		row = sheet.getRow(1);
		cell = row.getCell(2);
		if (cell == null) {
			cell = row.createCell(2);
		}

		cell.setCellValue("passed or failed");
	}

	@Test
	public void test2() {
		row = sheet.getRow(2);
		cell = row.getCell(2);
		if (cell == null) {
			cell = row.createCell(2);
		}
		cell.setCellValue("passed or failed");
	}

	@Test
	public void test3() {
		row = sheet.getRow(3);
		cell = row.getCell(2);
		if (cell == null) {
			cell = row.createCell(2);
		}
		cell.setCellValue("passed or failed");
	}

	@Test
	public void test4() {
		row = sheet.getRow(4);
		cell = row.getCell(2);
		if (cell == null) {
			cell = row.createCell(2);
		}
		cell.setCellValue("passed or failed");
	}

	@AfterClass
	public static void tearDown() throws IOException {

		FileOutputStream fileOut = new FileOutputStream(
				"C:\\Users\\Admin\\Desktop\\EclipseWorkspace\\Excel\\src\\main\\resources\\DemoSiteDDT.xlsx");

		workbook.write(fileOut);
		workbook.close();
		fileOut.close();

	}

}
