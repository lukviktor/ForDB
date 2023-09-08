package abbott.access;

import javax.swing.*;

public class DataAccess {
    private String inputFilePath() {
        return JOptionPane.showInputDialog(null, "Enter the path and name of the file DB Access\n" +
                "C:\\***\\***.***");
    }

    public static final String INPUT_FILE_PATH = new DataAccess().inputFilePath();
}
