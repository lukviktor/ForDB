package abbott.exel;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ChangingFiles {
    private void changingFiles(String inputFilePath, String outputFilePath) {
        try {
            // Открываем файл
            FileInputStream fis = new FileInputStream(inputFilePath);
            Workbook workbook1 = new XSSFWorkbook(fis);


            // Обработка изначального файла
            new ProcessWorkbook().processWorkbook(workbook1);


            // Записываем изменения в файл
            FileOutputStream fos = new FileOutputStream(outputFilePath);
            workbook1.write(fos);

            workbook1.close();
            fos.close();

            JOptionPane.showMessageDialog(null, "Обработка выполнена успешно!");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void convertFile() {
        DataExel dataExel = new DataExel();
        changingFiles(dataExel.inputFilePath(), DataExel.outputFilePath);
    }

}
