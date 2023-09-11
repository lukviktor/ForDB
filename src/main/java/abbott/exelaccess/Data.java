package abbott.exelaccess;

import javax.swing.*;

public class Data {
    private String inputFilePathExel() {
        return JOptionPane.showInputDialog(null, "Enter the path and name of the file exel to convert\n"
                + "C:\\***\\***.xlsx");
    }

    public static final String INPUT_FILE_PATH_EXEL = new Data().inputFilePathExel();

    private String outputFilePathExel() {
        String path = "C:\\Users\\";
        String name = JOptionPane.showInputDialog(null, "Enter the file name exel to save");
        return path + name + ".xlsx";
    }

    public static final String OUTPUT_FILE_PATH_EXEL = new Data().outputFilePathExel();

    private String inputFilePathAccess() {
        return JOptionPane.showInputDialog(null, "Enter the path and name of the file DB Access\n" +
                "C:\\***\\***.***");
    }

    public static final String INPUT_FILE_PATH_ACCESS = new Data().inputFilePathAccess();
}
