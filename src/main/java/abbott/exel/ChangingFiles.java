package abbott.exel;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ChangingFiles {
    public void changingFiles(String inputFilePath, String outputFilePath) {
        try {
            // Открываем файл
            FileInputStream fis = new FileInputStream(inputFilePath);
            Workbook workbook1 = new XSSFWorkbook(fis);


            // Обработка первого файла
            new ProcessWorkbook().processWorkbook(workbook1);


            // Записываем изменения в третий файл
            FileOutputStream fos = new FileOutputStream(outputFilePath);
            workbook1.write(fos);

            workbook1.close();
            fos.close();

            System.out.println("Обработка выполнена успешно!");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
