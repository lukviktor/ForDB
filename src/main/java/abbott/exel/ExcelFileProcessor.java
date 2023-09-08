package abbott.exel;

public class ExcelFileProcessor {
    public static void main(String[] args) {
        DataExel dataExel = new DataExel();
        new ChangingFiles().changingFiles(dataExel.getInputFilePath1(), dataExel.getOutputFilePath1());
        new ChangingFiles().changingFiles(dataExel.getInputFilePath2(), dataExel.getOutputFilePath2());
    }
}
