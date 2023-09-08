package abbott.exel;

import org.apache.poi.ss.usermodel.*;

public class ProcessWorkbook {
    void processWorkbook(Workbook workbook) {
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
