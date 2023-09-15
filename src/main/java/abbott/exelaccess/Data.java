package abbott.exelaccess;

import javax.swing.*;

public class Data {
    private String inputFilePathExel() {
        String massage = "Enter the path and name of the file exel to convert\n" +
                "Введите путь и имя файла excel для преобразования\n"
                + "Example/Пример -- C:\\***\\***.xlsx";

        return JOptionPane.showInputDialog(null, massage);
    }

    public static final String INPUT_FILE_PATH_EXEL = new Data().inputFilePathExel();

    private String outputFilePathExel() {
        String massage = "Enter the path and name of the file exel to save\n" +
                "Введите путь и имя файла excel для сохранения";

        String name = JOptionPane.showInputDialog(null, massage);
        return name + ".xlsx";
    }

    public static final String OUTPUT_FILE_PATH_EXEL = new Data().outputFilePathExel();

    private String inputFilePathAccess() {
        String massage = "Enter the path and name of the file DB Access\n" +
                "Введите путь и имя файла для доступа к базе данных\n" +
                "C:\\***\\***.***";

        return JOptionPane.showInputDialog(null, massage);
    }

    public static final String INPUT_FILE_PATH_ACCESS = new Data().inputFilePathAccess();
}
