package pages;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
    public void readExcel(String filePath,String fileName,String sheetName) throws IOException{
        File file =    new File(filePath+"/"+fileName);

        FileInputStream inputStream = new FileInputStream(file);
        Workbook guru99Workbook = null;
        String fileExtensionName = fileName.substring(fileName.indexOf("."));
        if(fileExtensionName.equals(".xlsx")){
            guru99Workbook = new XSSFWorkbook(inputStream);
        }
        else if(fileExtensionName.equals(".xls")){
            guru99Workbook = new HSSFWorkbook(inputStream);
        }
        Sheet guru99Sheet = guru99Workbook.getSheet(sheetName);
        int rowCount = guru99Sheet.getLastRowNum()-guru99Sheet.getFirstRowNum();
        System.out.println("row num:"+rowCount);

        Object[][] dataa = new Object[guru99Sheet.getLastRowNum()][guru99Sheet.getRow(0).getLastCellNum()];
        for (int i = 0; i < guru99Sheet.getLastRowNum()+1; i++) {
            Row row = guru99Sheet.getRow(i);
            for (int j = 0; j < row.getLastCellNum(); j++) {
                System.out.print(row.getCell(j)+ " ");
            }
            System.out.println();
        }

    }
    public static void main(String...strings) throws IOException{
       ReadExcel objExcelFile = new ReadExcel();
        String filePath = System.getProperty("user.dir")+"/testdata";
        objExcelFile.readExcel(filePath,"Book1.xlsx","Sheet1");
    }
}