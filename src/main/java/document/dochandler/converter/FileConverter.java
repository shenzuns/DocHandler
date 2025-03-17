package document.dochandler.converter;

import java.io.File;

public interface FileConverter {
    public void ToExcelConvert(File inputFile, String outputPath);
    public void ToWordConvert(File inputFile, String outputPath);
    public void ToPdfConvert(File inputFile, String outputPath);
    public void ToJsonConvert(File inputFile, String outputPath);
}
