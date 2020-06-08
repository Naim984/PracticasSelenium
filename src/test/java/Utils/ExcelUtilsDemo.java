package Utils;

import java.io.IOException;

public class ExcelUtilsDemo {
    public static void main(String[] args) throws IOException {
        String projectPath = System.getProperty("user.dir");
        ReadExcelFile excel = new ReadExcelFile(projectPath+"/excel/Libro12.xlsx", "Feuil1");
        excel.getRowCount();
        excel.getCellDataString(0,0);
        excel.getCellDataNumber(1,1);
    }
}
