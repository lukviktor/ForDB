package abbott.exel;

public class DataExel {
    private String inputFilePath1 = "src/main/resources/exel/generaList.xlsx";
    private String inputFilePath2 = "src/main/resources/exel/LAI.xlsx";
    private String outputFilePath1 = "src/main/resources/exel/outputLAI.xlsx";
    private String outputFilePath2 = "src/main/resources/exel/outputGeneraList.xlsx";

    public String getInputFilePath1() {
        return inputFilePath1;
    }

    public void setInputFilePath1(String inputFilePath1) {
        this.inputFilePath1 = inputFilePath1;
    }

    public String getInputFilePath2() {
        return inputFilePath2;
    }

    public void setInputFilePath2(String inputFilePath2) {
        this.inputFilePath2 = inputFilePath2;
    }

    public String getOutputFilePath1() {
        return outputFilePath1;
    }

    public void setOutputFilePath1(String outputFilePath1) {
        this.outputFilePath1 = outputFilePath1;
    }

    public String getOutputFilePath2() {
        return outputFilePath2;
    }

    public void setOutputFilePath2(String outputFilePath2) {
        this.outputFilePath2 = outputFilePath2;
    }
}
