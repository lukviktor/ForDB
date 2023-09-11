package newall;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ProcessingExcelFile {

    // Записываем изменения в файл
    public void writingChangesFile() {
        processWorkbook(openFileExel()); // Обработка изначального файла
    }

    // Открываем файл
    public Workbook openFileExel() {
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(Data.inputFilePathExel);
            Workbook workbook = new XSSFWorkbook(fis);
            return workbook;
        } catch (IOException e) {
            throw new RuntimeException(e);

        }
    }

    // Обработка изначального файла
    public void processExel() {
        processWorkbook(openFileExel());
    }

    // Обработка файла Exel
    void processWorkbook(Workbook workbook) {
        // обработка файла exel
        for (Sheet sheet : workbook) {
            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING) {
                        String cellValue = cell.getStringCellValue();

                        // Удаляем пробелы в конце строки
                        cellValue = cellValue.trim();

                        // Удаляем лишние пробелы внутри строки
                        cellValue = cellValue.replaceAll("\\s+", " ");

                        // Переводим строку в нижний регистр
                        cellValue = cellValue.toLowerCase();

                        cell.setCellValue(cellValue);
                    }
                }
            }
        }
    }
}
