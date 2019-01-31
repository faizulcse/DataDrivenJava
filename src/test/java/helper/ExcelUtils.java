package helper;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;

public class ExcelUtils {

	private static XSSFSheet ExcelWSheet;
	private static XSSFWorkbook ExcelWBook;
	private static XSSFCell xCell;
	private static XSSFRow xRow;

	// This method is to set the File path and to open the Excel file, Pass Excel
	// Path and Sheetname as Arguments to this method
	public static void setExcelFile(String Path, String SheetName) throws Exception {

		try {
			// Open the Excel file
			FileInputStream ExcelFile = new FileInputStream(Path);

			// Access the required test data sheet
			ExcelWBook = new XSSFWorkbook(ExcelFile);
			ExcelWSheet = ExcelWBook.getSheet(SheetName);
		} catch (Exception e) {
			throw (e);
		}

	}

	// This method is to read the test data from the Excel cell, in this we are
	// passing parameters as Row num and Col num
	public static String getCellData(int RowNum, int ColNum) throws Exception {

		try {
			xCell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
			String CellData = xCell.getStringCellValue();
			
			return CellData;
		} catch (Exception e) {
			return "";
		}

	}

	//This method is to write in the Excel cell, Row num and Col num are the parameters
	public static void setCellData(String Result, int RowNum, int ColNum) throws Exception {

		try {
			xRow = ExcelWSheet.getRow(RowNum);
			xCell = xRow.getCell(ColNum, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

			if (xCell == null) {
				xCell = xRow.createCell(ColNum);
				xCell.setCellValue(Result);
			} else {
				xCell.setCellValue(Result);
			}

			// Constant variables Test Data path and Test Data file name
			FileOutputStream fileOut = new FileOutputStream(AppConstant.PATH_TEST_DATA + AppConstant.FILE_TEST_DATA);
			ExcelWBook.write(fileOut);

			fileOut.flush();
			fileOut.close();

		} catch (Exception e) {
			throw (e);
		}

	}

}
