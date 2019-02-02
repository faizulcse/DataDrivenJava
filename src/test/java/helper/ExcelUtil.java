package helper;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;

public class ExcelUtil {

	private static XSSFSheet ExcelWSheet;
	private static XSSFWorkbook ExcelWBook;
	private static XSSFCell xCell;
	private static XSSFRow xRow;

	public static void setExcelFile(String Path, String SheetName) {
		try {
			ExcelWBook = new XSSFWorkbook(new FileInputStream(Path));
			ExcelWSheet = ExcelWBook.getSheet(SheetName);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static String getCellData(int RowNum, int ColNum) {
		try {
			xCell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
			return xCell.getStringCellValue();
		}
		catch (Exception e) {
			return "";
		}
	}

	public static void setCellData(String Result, int RowNum, int ColNum){
		try {
			xRow = ExcelWSheet.getRow(RowNum);
			xCell = xRow.getCell(ColNum, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

			if (xCell == null) {
				xCell = xRow.createCell(ColNum);
				xCell.setCellValue(Result);
			} else {
				xCell.setCellValue(Result);
			}

			FileOutputStream filePrint = new FileOutputStream(AppConstant.PATH_TEST_DATA + AppConstant.FILE_TEST_DATA);
			ExcelWBook.write(filePrint);
			filePrint.flush();
			filePrint.close();
		}
		catch (Exception e) {
			e.printStackTrace();
		}
	}
}
