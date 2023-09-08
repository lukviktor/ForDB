package abbott.exel;

import javax.swing.*;

public class DataExel {
    private String inputFilePath() {
        return JOptionPane.showInputDialog(null, "Enter the path and name of the file exel to convert\n" +
                "C:\\***\\***.xlsx");
    }

    private String outputFilePath() {
        String path = JOptionPane.showInputDialog(null, "Enter the path to the file exel to save\n" +
                "C:\\***\\***");
        String name = JOptionPane.showInputDialog(null, "Enter the file name exel to save");
        return path + name + ".xlsx";
    }

    public static final String INPUT_FILE_PATH = new DataExel().inputFilePath();

    public static final String OUTPUT_FILE_PATH = new DataExel().outputFilePath();
}
