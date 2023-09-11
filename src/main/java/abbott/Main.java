package abbott;

import abbott.exelaccess.ExcelToAccess;
import abbott.exelaccess.ProcessingExcelFile;

public class Main {

    public static void main(String[] args) {
        new ProcessingExcelFile().openingAndModifyingExelFile();
        new ExcelToAccess().excelToAccess();
    }
}
