package abbott.exelaccess;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class ProcessingExcelFile {
    // Создаем поток
    private FileInputStream fis() {
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(Data.INPUT_FILE_PATH_EXEL);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
        return fis;
    }

    // Открываем файл
    private Workbook workbook() {
        Workbook workbook = null;
        try {
            workbook = new XSSFWorkbook(fis());
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return workbook;
    }

    // Обработка изначального файла и запись изменения в файл
    private void processExel() {
        new ProcessingExcelFile().processWorkbookExel(workbook());
        try {
            FileOutputStream fos = new FileOutputStream(Data.OUTPUT_FILE_PATH_EXEL);
            workbook().write(fos);
            workbook().close();
            fos.close();
            fis().close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Обработка файла Exel
    private void processWorkbookExel(Workbook workbook) {
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

    public void openingAndModifyingExelFile() {
        processExel();
    }
}
