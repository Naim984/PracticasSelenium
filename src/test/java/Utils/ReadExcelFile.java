package Utils;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

public class ReadExcelFile {
    static String projectPath;
    static XSSFWorkbook workbook;
    static XSSFSheet sheet;

    public ReadExcelFile(String excelPath, String sheeName) throws IOException {
        try {
         //   projectPath = System.getProperty("user.dir");
            workbook = new XSSFWorkbook(excelPath);
            sheet = workbook.getSheet("Feuil1");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    public static void main(String[] args) {
        //getRowCount();
      //  getCellDataString(0,0);
       // getCellDataNumber(1,1);
    }

    public static void getRowCount() {
 try {
        projectPath = System.getProperty("user.dir");
        workbook = new XSSFWorkbook(projectPath+"/excel/Libro12.xlsx");
        sheet = workbook.getSheet("Feuil1");
        int rowCount = sheet.getPhysicalNumberOfRows();
        System.out.println("No of rows : " + rowCount);


    }catch (Exception exp){
     System.out.println(exp.getMessage());
     System.out.println(exp.getCause());
     exp.printStackTrace();
 }

 }

    public static void getCellDataString(int rowNum, int colNum) {
        try {
          //  projectPath = System.getProperty("user.dir");
          // workbook = new XSSFWorkbook(projectPath + "/excel/Libro12.xlsx");
          //  workbook = new XSSFWorkbook(projectPath + "/excel/ejemplo.xlsx");
          //  sheet = workbook.getSheet("Feuil1");
           String CellData = sheet.getRow(rowNum).getCell(colNum).getStringCellValue();
           System.out.println(CellData);
        } catch (Exception exp) {
            System.out.println(exp.getMessage());
            System.out.println(exp.getCause());
            exp.printStackTrace();

        }

    }
    public static void getCellDataNumber(int rowNum, int colNum) {
        try {
            //projectPath = System.getProperty("user.dir");
            //workbook = new XSSFWorkbook(projectPath + "/excel/Libro12.xlsx");
            //  workbook = new XSSFWorkbook(projectPath + "/excel/ejemplo.xlsx");
            // sheet = workbook.getSheet("Feuil1");
          double cellData = sheet.getRow(rowNum).getCell(colNum).getNumericCellValue();
            System.out.println(cellData);
        } catch (Exception exp) {
            System.out.println(exp.getMessage());
            System.out.println(exp.getCause());
            exp.printStackTrace();

        }

        }
    }


