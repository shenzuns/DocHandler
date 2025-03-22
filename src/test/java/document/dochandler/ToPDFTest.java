package document.dochandler;

import document.dochandler.config.DocConfigLoader;
import document.dochandler.converter.impl.ToPDFConverter;
import document.dochandler.exception.FileConverterException;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;

import static org.assertj.core.api.Fail.fail;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class ToPDFTest {
    String rootPath = "./documents/";
    public static void main(String[] args) {
        DocConfigLoader configLoader = new DocConfigLoader();

        File inputFile = new File("documents/sxs.txt");
        String outputPath = "documents/sxs.pdf";
        ToPDFConverter converter = new ToPDFConverter(configLoader);
        try {
            // 调用转换方法
            File convertedFile = converter.ToPdfConvert(inputFile, outputPath);

            // 验证文件存在
            assertTrue(convertedFile.exists(), "转换后的文件不存在：" + convertedFile.getPath());

            // 日志输出
            System.out.println("成功将文件转换为 PDF：" + inputFile.getName());
        } catch (FileConverterException e) {
            // 捕获异常并打印错误信息
            fail("文件转换失败：" + inputFile.getName() + "，错误原因：" + e.getMessage());
        }

    }
    @Test
    public void toPdfTest() {
        DocConfigLoader configLoader = new DocConfigLoader();
        ToPDFConverter converter = new ToPDFConverter(configLoader);

        String[] testFiles = {
                "sxs.txt",
                "智控院反诈教育安排.xlsx",
                "补考复习提纲.docx"
        };

        for (String fileName : testFiles) {
            File inputFile = new File("./documents/" + fileName );
            String outputPath = rootPath + fileName.replaceFirst("[.][^.]+$", "") + "_converted.pdf";

            try {
                // 调用转换方法
                File convertedFile = converter.ToPdfConvert(inputFile, outputPath);

                // 验证文件存在
                assertTrue(convertedFile.exists(), "转换后的文件不存在：" + convertedFile.getPath());

                // 日志输出
                System.out.println("成功将文件转换为 PDF：" + inputFile.getName());
            } catch (FileConverterException e) {
                // 捕获异常并打印错误信息
                fail("文件转换失败：" + inputFile.getName() + "，错误原因：" + e.getMessage());
                }
        }
    }

}
