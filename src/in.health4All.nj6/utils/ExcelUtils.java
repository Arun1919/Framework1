import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelUtils {

    private static XSSFSheet excelWSheet;
    private static XSSFWorkbook excelWBook;
    private static XSSFCell cell;
    private static XSSFRow row;

    public static void setExcelFile(String path,String sheetName) throws Exception {
        try {
            FileInputStream excelFile = new FileInputStream(path);
            excelWBook = new XSSFWorkbook(ExcelFile);
            excelWSheet = excelWBook.getSheet(sheetName);

            }
        catch (Exception e) {
            throw (e);
            }
    }

     public static String getCellData(int rowNum,int colNum) throws Exception {

         try {
             cell = excelWSheet.getRow(rowNum).getCell(colNum);
             String cellData = cell.getStringCellValue();
             return cellData;

         } catch (Exception e) {
             throw (e);

         }

     }
    //To write in the Excel Cell
     public static void setCellData(String result, int rowNum, int colNum) throws Exception {

         try {
             row = excelWSheet.getRow(rowNum);
             cell = row.getCell(colNum, Row.RETURN_BLANK_AS_NULL);
             if (cell == null) {
                 cell = row.createCell(colNum);
                 cell.setCellValue(result);
             } else {
                 Cell.setCellValue(result);
             }

             FileOutputStream fileOut = new FileOutputStream(Constants.Path_TestData + Constants.File_TestData);
             ExcelWBook.write(fileOut);
             fileOut.flush();
             fileOut.close();
         } catch (Exception e) {
             throw (e);
         }
     }

     }





