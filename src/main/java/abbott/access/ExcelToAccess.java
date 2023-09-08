package abbott.access;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToAccess {

    public void excelToAccess() {

    }
    public static void main(String[] args) {
        String excelFilePath = "path/to/your/file.xlsx";

        try (Connection connection = DriverManager.getConnection("jdbc:ucanaccess://path/to/your/accessdatabase.accdb")) {

            Workbook workbook = XSSFWorkbook(excelFilePath);
            int numberOfSheets = workbook.getNumberOfSheets();

            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String tableName = sheet.getSheet
                // Создание таблицы в базе данных Access
                createTable(connection, sheet, tableName);

                // Внесение данных в таблицу
                insertData(connection, sheet, tableName);
            }

            workbook.close();
            System.out.println("Данные успешно импортированы в базу данных Access!");
        } catch (SQLException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void createTable(Connection connection, Sheet sheet, String tableName) throws SQLException {
        StringBuilder createTableQuery = new StringBuilder("CREATE TABLE ");
        createTableQuery.append(tableName).append(" (");

        Row headerRow = sheet.getRow(0);
        Iterator<Cell> cellIterator = headerRow.cellIterator();
        while (.hasNext()) {
            Cell cell = cellIterator.next();
            createTableQuery
                    .append(cell.getStringCellValue())
                    .append(" VARCHAR(255)");
            if (cellIterator.hasNext()) {
                createTableQuery.append(", ");
            }
        }

        createTableQuery.append(")");

        PreparedStatement createTableStatement = connection.prepareStatement(createTableQuery.toString());
        createTableStatement.executeUpdate();
        createTableStatement.close();
    }

    private static void insertData(Connection connection, Sheet sheet, String tableName) throws SQLException {
        String insertQuery = "INSERT INTO " + tableName + " VALUES (";

        Iterator<Row> rowIterator = sheet.iterator();
        // Пропускаем первую строку, так как это заголовки
        rowIterator.next();

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();

            StringBuilder rowData = new StringBuilder(insertQuery);
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                rowData
                        .append("'")
                        .append(cell.getStringCellValue())
                        .append("'");
                if (cellIterator.hasNext()) {
                    rowData.append(", ");
                }
            }
            rowData.append(")");

            PreparedStatement insertStatement = connection.prepareStatement(rowData.toString());
            insertStatement.executeUpdate();
            insertStatement.close();
        }
    }
}
