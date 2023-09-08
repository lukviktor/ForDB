package abbott.exel;

import javax.swing.*;

public class DataExel {
    public String inputFilePath() {
        return JOptionPane.showInputDialog(null, "Enter the path and name of the file to convert\n" +
                "C:\\***\\***\\***.xlsx");
    }

    public String outputFilePath() {
        String path = JOptionPane.showInputDialog(null, "Enter the path to the file to save\n" +
                "C:\\***\\***");
        String name = JOptionPane.showInputDialog(null, "Enter the file name to save");
        return path + name + ".xlsx";
    }
}
