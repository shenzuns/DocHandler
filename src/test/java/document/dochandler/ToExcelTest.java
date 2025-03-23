package document.dochandler;

import document.dochandler.config.DocConfigLoader;
import document.dochandler.converter.impl.ToExcelConverter;
import document.dochandler.exception.FileConverterException;
import org.junit.jupiter.api.Test;

import java.io.File;

import static org.assertj.core.api.Assertions.fail;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class ToExcelTest {
    String rootPath = "./documents/";

    @Test
    public void toExcelTest() {

        ToExcelConverter converter = new ToExcelConverter();

        String[] testFiles = {
                "sxs.txt",
                "补考复习提纲.docx"
        };

        for (String fileName : testFiles) {
            File inputFile = new File("./documents/" + fileName);
            String outputPath = rootPath + fileName.replaceFirst("[.][^.]+$", "") + "_converted.xlsx";

            try {
                // 调用转换方法
                File convertedFile = converter.ToExcelConvert(inputFile, outputPath);

                // 验证文件是否存在
                assertTrue(convertedFile.exists(), "转换后的文件不存在：" + convertedFile.getPath());

                // 日志输出
                System.out.println("成功将文件转换为 Excel：" + inputFile.getName());
            } catch (FileConverterException e) {
                // 捕获异常并打印错误信息
                fail("文件转换失败：" + inputFile.getName() + "，错误原因：" + e.getMessage());
            }
        }
    }
}
