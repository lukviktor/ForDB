package abbott.exel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelFileProcessor {
    public static void main(String[] args) {
        DataExel dataExel = new DataExel();
        new ChangingFiles().changingFiles(dataExel.getInputFilePath1(), dataExel.getOutputFilePath1());
        new ChangingFiles().changingFiles(dataExel.getInputFilePath2(), dataExel.getOutputFilePath2());
    }
}
