package helper;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TestUtil {
	
	public static String TESTDATA_SHEET_PATH = "testdata/Book1.xlsx";
	static Workbook book;
	static Sheet sheet;
	
	public static Object[][] getTestData(String sheetName) {
		System.out.println(sheetName);
		try {
			book = new XSSFWorkbook(new FileInputStream(TESTDATA_SHEET_PATH));
			sheet = book.getSheet(sheetName);
		}
		catch (IOException e) {
			e.printStackTrace();
		}


		//System.out.println(sheet.getLastRowNum() + "--------" + sheet.getRow(0).getLastCellNum());
		Object[][] dataa = new Object[sheet.getLastRowNum()][sheet.getRow(0).getLastCellNum()];
		for (int i = 0; i < sheet.getLastRowNum(); i++) {
			int row = sheet.getRow(i).getLastCellNum();
			//System.out.println(row);
			for (int k = 0; k < row; k++) {
				dataa[i][k] = sheet.getRow(i).getCell(k).toString();//row.getCell(k).getStringCellValue();
				 System.out.print(dataa[i][k]+ " || ");
			}
			System.out.println();
		}

		Object[][] data = new Object[sheet.getLastRowNum()][sheet.getRow(0).getLastCellNum()];
		return data;
	}

}
