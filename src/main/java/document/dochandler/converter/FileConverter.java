package document.dochandler.converter;

import java.io.File;

public interface FileConverter {
    public File ToExcelConvert(File inputFile, String outputPath);
    public File ToWordConvert(File inputFile, String outputPath);
    public File ToPdfConvert(File inputFile, String outputPath);
    public File ToJsonConvert(File inputFile, String outputPath);
}
