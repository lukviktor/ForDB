package abbott.access;

import abbott.exel.DataExel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.Iterator;

public class ExcelToAccess {

    public void excelToAccess() {
        String excelFilePath = DataExel.INPUT_FILE_PATH;
        String databasePath = DataAccess.INPUT_FILE_PATH;

        try (Connection connection = DriverManager.getConnection("jdbc:ucanaccess://" + databasePath)) {
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(fis);
            int numberOfSheets = workbook.getNumberOfSheets();

            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String tableName = sheet.getSheetName();
                // Создание таблицы в базе данных Access
                createTable(connection, sheet, tableName);

                // Внесение данных в таблицу
                insertData(connection, sheet, tableName);
            }

            workbook.close();
            fis.close();
            System.out.println("Данные успешно импортированы в базу данных Access!");
        } catch (SQLException | IOException e) {
            throw new RuntimeException(e);
        }
    }


    private static void createTable(Connection connection, Sheet sheet, String tableName) {
        StringBuilder createTableQuery = new StringBuilder("CREATE TABLE ");
        createTableQuery.append(tableName).append(" (");

        Row headerRow = sheet.getRow(0);
        Iterator<Cell> cellIterator = headerRow.cellIterator();
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            createTableQuery
                    .append(cell.getStringCellValue())
                    .append(" VARCHAR(255)");
            if (cellIterator.hasNext()) {
                createTableQuery.append(", ");
            }
        }

        createTableQuery.append(")");

        PreparedStatement createTableStatement = null;
        try {
            createTableStatement = connection.prepareStatement(createTableQuery.toString());
        } catch (SQLException e) {
            throw new RuntimeException(e);
        }
        try {
            createTableStatement.executeUpdate();
        } catch (SQLException e) {
            throw new RuntimeException(e);
        }
        try {
            createTableStatement.close();
        } catch (SQLException e) {
            throw new RuntimeException(e);
        }
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

    public static void main(String[] args) {
        //new ExcelToAccess().excelToAccess();
        System.out.println(DataExel.INPUT_FILE_PATH);
    }
}
