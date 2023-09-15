package abbott.testaccess;
import java.io.File;
import java.nio.file.Files;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;

public class AccessExample {
    public static void main(String[] args) {
        try {
            // Путь до файла базы данных Access
            String databasePath = "src/main/resources/access/testDB.accdb";

            // Копируем файл базы данных во временный файл
            File tempFile = File.createTempFile("temp", ".accdb");
            tempFile.deleteOnExit();
            Files.copy(AccessExample.class.getResourceAsStream("/" + databasePath), tempFile.toPath());

            // Загружаем JDBC драйвер UCanAccess
//            String driverClass = "net.ucanaccess.jdbc.UcanaccessDriver";
//            URLClassLoader.loadClass(driverClass);

            // Устанавливаем соединение с базой данных Access
            String connectionString = "jdbc:ucanaccess://" + tempFile.getAbsolutePath();
            Connection conn = DriverManager.getConnection(connectionString);

            // Создаем объект для выполнения SQL-запросов
            Statement stmt = conn.createStatement();

            // Выполняем запрос
            ResultSet rs = stmt.executeQuery("SELECT * FROM your_table");

            // Обрабатываем результаты
            while (rs.next()) {
                System.out.println(rs.getString("column_name"));
            }

            // Закрываем соединение
            rs.close();
            stmt.close();
            conn.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
