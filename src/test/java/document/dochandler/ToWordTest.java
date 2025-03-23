package document.dochandler;

import document.dochandler.config.DocConfigLoader;
import document.dochandler.converter.impl.ToPDFConverter;
import document.dochandler.converter.impl.ToWordConverter;
import document.dochandler.exception.FileConverterException;
import org.junit.jupiter.api.Test;

import java.io.File;

import static org.assertj.core.api.Fail.fail;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class ToWordTest {
    String rootPath = "./documents/";
    @Test
    public void toWordTest() {
        DocConfigLoader configLoader = new DocConfigLoader();
        ToWordConverter converter = new ToWordConverter(configLoader);

        String[] testFiles = {
                "智控院反诈教育安排.xlsx",
                "sxs.txt",
                "2024奖学金推荐获奖人.pdf"
        };

//        for (String fileName : testFiles) {
            String fileName = testFiles[2];
            File inputFile = new File("./documents/" + testFiles[2] );
            String outputPath = rootPath + fileName.replaceFirst("[.][^.]+$", "") + "_converted.docx";

            try {
                // 调用转换方法
                File convertedFile = converter.ToWordConvert(inputFile, outputPath);

                // 验证文件存在
                assertTrue(convertedFile.exists(), "转换后的文件不存在：" + convertedFile.getPath());

                // 日志输出
                System.out.println("成功将文件转换为 Word：" + inputFile.getName());
            } catch (FileConverterException e) {
                // 捕获异常并打印错误信息
                fail("文件转换失败：" + inputFile.getName() + "，错误原因：" + e.getMessage());
//            }
        }
    }
}
