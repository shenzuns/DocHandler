package document.dochandler.converter;

import document.dochandler.config.DocConfigLoader;
import document.dochandler.exception.BaseException;
import document.dochandler.exception.FileConverterException;
import document.dochandler.utils.FileValidatorUtils;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import java.util.concurrent.atomic.AtomicReference;

import java.io.*;

@Component
public class ToPDFConverter {
    private final DocConfigLoader configLoader;
    private final String font;
    private final float fontSize;
    private final float marginTop;
    private final float marginLeft;
    private final float initialY;
    private final float lineSpacing;

    @Autowired
    public ToPDFConverter(DocConfigLoader configLoader) {
        this.configLoader = configLoader;
        this.font = configLoader.getPdfFont();
        this.fontSize = configLoader.getPdfFontSize();
        this.marginTop = configLoader.getPdfMarginTop();
        this.marginLeft = configLoader.getPdfMarginLeft();
        this.initialY = 800 - marginTop;
        this.lineSpacing = configLoader.getPdfLineSpacing();
    }
    /**
     * 将文件转换为 PDF 类型
     *
     * @param inputFile  输入文件
     * @param outputPath 输出文件路径
     * @return 转换后的 PDF 文件
     * @throws FileConverterException 如果转换失败
     */
    public File convertTxtToPdf(File inputFile, String outputPath) {
        PDDocument document = new PDDocument();

        try (BufferedReader reader = new BufferedReader(new FileReader(inputFile))) {
            PDPage page = new PDPage();
            document.addPage(page);

            // 初始化字体和页面流
            InputStream fontStream = new FileInputStream("src/main/resources/fonts/" + font);
            PDFont chineseFont = PDType0Font.load(document, fontStream);
            PDPageContentStream contentStream = new PDPageContentStream(document, page);
            contentStream.setFont(chineseFont, fontSize);

            float pageWidth = page.getMediaBox().getWidth();
            float marginRight = pageWidth - marginLeft;
            AtomicReference<Float> yRef = new AtomicReference<>(initialY); // 使用AtomicReference管理y

            String line;
            while ((line = reader.readLine()) != null) {
                contentStream = wordArrange(
                        line.toCharArray(), document, chineseFont, marginRight, yRef, page, contentStream);
            }

            contentStream.close();
            document.save(outputPath); // 保存PDF文档
        } catch (IOException e) {
            throw new FileConverterException("TXT 转 PDF 失败：" + e.getMessage());
        } finally {
            try {
                document.close();
            } catch (IOException e) {
                throw new FileConverterException("关闭 PDF 失败：" + e.getMessage());
            }
        }
        return new File(outputPath);
    }

    private File convertWordToPdf(File inputFile, String outputPath) {
        PDDocument document = new PDDocument();
        PDPageContentStream contentStream = null;

        try (FileInputStream inputStream = new FileInputStream(inputFile)) {
            // 加载字体
            InputStream fontStream = new FileInputStream("src/main/resources/fonts/" + font);
            PDFont configuredFont = PDType0Font.load(document, fontStream);

            XWPFDocument wordDoc = new XWPFDocument(inputStream);

            PDPage page = new PDPage();
            document.addPage(page);
            contentStream = createContentStream(document, page, configuredFont);

            // 页面尺寸相关配置
            float pageWidth = page.getMediaBox().getWidth();
            float marginRight = pageWidth - marginLeft;  // 右边界
            AtomicReference<Float> yRef = new AtomicReference<>(initialY); // 使用AtomicReference管理y

            for (XWPFParagraph paragraph : wordDoc.getParagraphs()) {
                String text = paragraph.getText();
                if (!text.isEmpty()) {
                    contentStream = wordArrange(
                            text.toCharArray(), document, configuredFont, marginRight, yRef, page, contentStream);
                }
            }

            contentStream.close();
            document.save(outputPath);
        } catch (IOException e) {
            throw new FileConverterException("Word 转 PDF 失败：" + e.getMessage());
        } finally {
            try {
                document.close();
            } catch (IOException e) {
                throw new FileConverterException("保存 PDF 失败：" + e.getMessage());
            }
        }
        return new File(outputPath);
    }

    private File convertExcelToPdf(File inputFile, String outputPath) {
        PDDocument document = new PDDocument();
        PDPageContentStream contentStream = null;

        try (Workbook workbook = WorkbookFactory.create(inputFile)) {
            // 加载字体
            InputStream fontStream = new FileInputStream("src/main/resources/fonts/" + font);
            PDFont configuredFont = PDType0Font.load(document, fontStream);

            Sheet sheet = workbook.getSheetAt(0);

            PDPage page = new PDPage();
            document.addPage(page);
            contentStream = createContentStream(document, page, configuredFont);

            // 页面尺寸相关配置
            float pageWidth = page.getMediaBox().getWidth();
            float marginRight = pageWidth - marginLeft;  // 右边界
            AtomicReference<Float> yRef = new AtomicReference<>(initialY); // 使用AtomicReference管理y

            for (Row row : sheet) {
                StringBuilder rowText = new StringBuilder();

                for (Cell cell : row) {
                    String cellText = cell.toString().replace("\t", "    "); // 替换 '\t' 为四个空格
                    rowText.append(cellText).append(" "); // 单元格分隔
                }

                float lineWidth = marginLeft; // 当前行宽，初始为左边距
                contentStream = wordArrange(
                        rowText.toString().toCharArray(), document, configuredFont, marginRight, yRef, page, contentStream);
            }

            contentStream.close();
            document.save(outputPath);
        } catch (IOException e) {
            throw new FileConverterException("Excel 转 PDF 失败：" + e.getMessage());
        } finally {
            try {
                document.close();
            } catch (IOException e) {
                throw new FileConverterException("保存 PDF 失败：" + e.getMessage());
            }
        }
        return new File(outputPath);
    }

    private PDPageContentStream createContentStream(PDDocument document, PDPage page, PDFont font) throws IOException {
        PDPageContentStream contentStream = new PDPageContentStream(document, page);
        contentStream.setFont(font, fontSize);
        return contentStream;
    }

    private PDPageContentStream wordArrange(
            char[] array,
            PDDocument document,
            PDFont configuredFont,
            float marginRight,
            AtomicReference<Float> yRef,
            PDPage page,
            PDPageContentStream contentStream) throws IOException {

        float y = yRef.get(); // 解包当前的y
        float lineWidth = marginLeft; // 当前行宽，初始为左边距

        for (char c : array) {
            float charWidth = configuredFont.getStringWidth(String.valueOf(c)) / 1000 * fontSize;

            // 如果字符宽度超出右边界，则换行
            if (lineWidth + charWidth > marginRight) {
                y -= (fontSize + lineSpacing); // 行间距
                lineWidth = marginLeft; // 重置行宽

                // 如果Y坐标超出页面底部，换页
                if (y < marginTop) {
                    contentStream.close();
                    page = new PDPage();
                    document.addPage(page);
                    y = initialY; // 重置Y坐标
                    contentStream = new PDPageContentStream(document, page);
                    contentStream.setFont(configuredFont, fontSize);
                }
            }

            // 写入字符
            contentStream.beginText();
            contentStream.newLineAtOffset(lineWidth, y);
            contentStream.showText(String.valueOf(c));
            contentStream.endText();

            lineWidth += charWidth; // 更新行宽
        }

        // 更新并存储最新的y
        y -= (fontSize + lineSpacing);
        yRef.set(y); // 更新y的全局引用

        // 换页逻辑
        if (y < marginTop) {
            contentStream.close();
            page = new PDPage();
            document.addPage(page);
            y = initialY;
            contentStream = new PDPageContentStream(document, page);
            contentStream.setFont(configuredFont, fontSize);
            yRef.set(y); // 再次更新y
        }

        return contentStream; // 返回最新的contentStream
    }

}
