package abbott.testaccess;


import java.sql.*;

/**
 * This program demonstrates how to use UCanAccess JDBC driver to read/write
 * a Microsoft Access database.
 *
 */
public class JdbcAccessTest {

    public static void main(String[] args) {

        String databaseURL = "jdbc:ucanaccess://e://Java//JavaSE//MsAccess//Contacts.accdb";

        try (Connection connection = DriverManager.getConnection(databaseURL)) {


            String sql = "INSERT INTO Contacts (Full_Name, Email, Phone) VALUES (?, ?, ?)";

            PreparedStatement preparedStatement = connection.prepareStatement(sql);
            preparedStatement.setString(1, "Jim Rohn");
            preparedStatement.setString(2, "rohnj@herbalife.com");
            preparedStatement.setString(3, "0919989998");

            int row = preparedStatement.executeUpdate();

            if (row > 0) {
                System.out.println("A row has been inserted successfully.");
            }

            sql = "SELECT * FROM Contacts";

            Statement statement = connection.createStatement();
            ResultSet result = statement.executeQuery(sql);

            while (result.next()) {
                int id = result.getInt("Contact_ID");
                String fullname = result.getString("Full_Name");
                String email = result.getString("Email");
                String phone = result.getString("Phone");

                System.out.println(id + ", " + fullname + ", " + email + ", " + phone);
            }

        } catch (SQLException ex) {
            ex.printStackTrace();
        }
    }
}