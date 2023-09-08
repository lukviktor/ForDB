package abbott.exel;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class ChangingFiles {

    private void changingFiles(String inputFilePath, String outputFilePath) {
        try {
            // Открываем файл
            FileInputStream fis = new FileInputStream(inputFilePath);
            Workbook workbook = new XSSFWorkbook(fis);


            // Обработка изначального файла
            new ProcessWorkbook().processWorkbook(workbook);


            // Записываем изменения в файл
            FileOutputStream fos = new FileOutputStream(outputFilePath);
            workbook.write(fos);

            workbook.close();
            fos.close();
            fis.close();

            JOptionPane.showMessageDialog(null, "Обработка выполнена успешно!");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void convertFile() {
        DataExel dataExel = new DataExel();
        changingFiles(DataExel.INPUT_FILE_PATH, DataExel.OUTPUT_FILE_PATH);
    }

}
