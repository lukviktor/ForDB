package abbott.work;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.Statement;

public class ExcelToAccessImporter {
    public static void main(String[] args) {

        String excelFilePath = "ex.xlsx";
        String accessFilePath = "src/main/resources/db.accdb";

        try {
            // Load JDBC драйвера for Access
            Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");

            // Connection Data Base Access
            Connection conn = DriverManager.getConnection("jdbc:ucanaccess://" + accessFilePath);

            // Read Excel
            Workbook workbook = new XSSFWorkbook(new File(excelFilePath));

            // read every list in excels
            for (Sheet sheet : workbook) {
                String tableName = sheet.getSheetName();

                // Create table Data Base Access
                String createTableQuery = "CREATE TABLE " + tableName + " (";
                Row headerRow = sheet.getRow(0);
                for (Cell cell : headerRow) {
                    String columnName = cell.getStringCellValue();
                    createTableQuery += columnName + " TEXT, ";
                }
                createTableQuery = createTableQuery.substring(0, createTableQuery.length() - 2) + ")";
                Statement statement = conn.createStatement();
                statement.executeUpdate(createTableQuery);

                // Insert data in table
                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row dataRow = sheet.getRow(i);
                    String insertQuery = "INSERT INTO " + tableName + " VALUES (";
                    for (int j = 0; j < dataRow.getLastCellNum(); j++) {
                        Cell dataCell = dataRow.getCell(j);
                        if (dataCell.getCellType() == CellType.NUMERIC) {
                            insertQuery += dataCell.getNumericCellValue() + ", ";
                        } else {
                            insertQuery += "'" + dataCell.getStringCellValue() + "', ";
                        }
                    }
                    insertQuery = insertQuery.substring(0, insertQuery.length() - 2) + ")";
                    statement.executeUpdate(insertQuery);
                }
            }

            // Close recurse
            workbook.close();
            conn.close();


            System.out.println("Импорт данных успешно выполнен.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
