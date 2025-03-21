package document.dochandler.converter.impl;

import document.dochandler.config.DocConfigLoader;
import document.dochandler.converter.FileConverter;
import document.dochandler.exception.BaseException;
import document.dochandler.exception.FileConverterException;
import document.dochandler.utils.FileValidatorUtils;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.io.*;
import java.util.List;
@Component
public class ToPDFConverter implements FileConverter {
    private final DocConfigLoader configLoader;
    private final float fontSize;
    private final float marginTop;
    private final float marginLeft;
    private final float initialY;
    private final float lineSpacing;

    @Autowired
    public ToPDFConverter(DocConfigLoader configLoader) {
        this.configLoader = configLoader;
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
    @Override
    public File ToPdfConvert(File inputFile, String outputPath) {
        try {
            if (!FileValidatorUtils.isFileValid(inputFile)) {
                throw new FileConverterException("输入文件无效");
            }

            if (outputPath == null || outputPath.isEmpty()) {
                outputPath = inputFile.getParent() + File.separator +
                        inputFile.getName().replaceFirst("[.][^.]+$", "") + ".pdf";
            }

            if (inputFile.getName().endsWith(".txt")) {
                return convertTxtToPdf(inputFile, outputPath);
            } else if (inputFile.getName().endsWith(".xlsx") || inputFile.getName().endsWith(".xls")) {
                return convertExcelToPdf(inputFile, outputPath);
            } else if (inputFile.getName().endsWith(".doc") || inputFile.getName().endsWith(".docx")) {
                return convertWordToPdf(inputFile, outputPath);
            } else {
                throw new FileConverterException("不支持的文件类型：" + inputFile.getName());
            }
        } catch (Exception e) {
            throw new FileConverterException("文件转换失败：" + e.getMessage());
        }
    }

    @Override
    public File ToJsonConvert(File inputFile, String outputPath) {
        throw new BaseException("该实现类仅支持转PDF");
    }
    @Override
    public File ToExcelConvert(File inputFile, String outputPath) {
        throw new BaseException("该实现类仅支持转PDF");
    }
    @Override
    public File ToWordConvert(File inputFile, String outputPath) {
        throw new BaseException("该实现类仅支持转PDF");
    }

    private File convertTxtToPdf(File inputFile, String outputPath) {
        PDDocument document = new PDDocument();
        PDPageContentStream contentStream = null;

        try (BufferedReader reader = new BufferedReader(new FileReader(inputFile))) {
            PDPage page = new PDPage();
            document.addPage(page);

            // 加载中文字体
            InputStream fontStream = new FileInputStream("src/main/resources/fonts/simhei.ttf");
            PDFont chineseFont = PDType0Font.load(document, fontStream);
            contentStream = new PDPageContentStream(document, page);
            contentStream.setFont(chineseFont, fontSize);

            // 页面尺寸相关配置
            float pageWidth = page.getMediaBox().getWidth();
            float marginRight = pageWidth - marginLeft;  // 右边界
            float y = initialY;  // 起始Y坐标

            String line;
            while ((line = reader.readLine()) != null) {
                float lineWidth = marginLeft; // 当前行宽，初始为左边距

                for (char c : line.toCharArray()) {
                    float charWidth = chineseFont.getStringWidth(String.valueOf(c)) / 1000 * fontSize;

                    // 如果字符宽度超出右边界，则换行
                    if (lineWidth + charWidth > marginRight) {
                        y -= (fontSize + 3); // 行间距
                        lineWidth = marginLeft;  // 重置行宽

                        // 如果Y坐标超出页面底部，换页
                        if (y < marginTop) {
                            contentStream.close();
                            page = new PDPage();
                            document.addPage(page);
                            y = initialY; // 重置Y坐标
                            contentStream = new PDPageContentStream(document, page);
                            contentStream.setFont(chineseFont, fontSize);
                        }
                    }

                    // 写入字符
                    contentStream.beginText();
                    contentStream.newLineAtOffset(lineWidth, y);
                    contentStream.showText(String.valueOf(c));
                    contentStream.endText();

                    lineWidth += charWidth; // 更新行宽
                }

                // 行结束后处理行间距
                y -= (fontSize + 3);

                // 如果Y坐标超出页面底部，换页
                if (y < marginTop) {
                    contentStream.close();
                    page = new PDPage();
                    document.addPage(page);
                    y = initialY;  // 重置Y坐标
                    contentStream = new PDPageContentStream(document, page);
                    contentStream.setFont(chineseFont, fontSize);
                }
            }

            contentStream.close();
        } catch (IOException e) {
            throw new FileConverterException("TXT 转 PDF 失败：" + e.getMessage());
        } finally {
            try {
                document.save(outputPath);
                document.close();
            } catch (IOException e) {
                throw new FileConverterException("保存 PDF 失败：" + e.getMessage());
            }
        }
        return new File(outputPath);
    }

    private File convertWordToPdf(File inputFile, String outputPath) {
        try (PDDocument document = new PDDocument();
             FileInputStream inputStream = new FileInputStream(inputFile)) {
            XWPFDocument wordDoc = new XWPFDocument(inputStream);
            PDPage page = new PDPage();
            document.addPage(page);
            PDPageContentStream contentStream = new PDPageContentStream(document, page);

            try {
                // 设置字体
                InputStream fontStream = new FileInputStream("src/main/resources/fonts/simshei.ttf");
                PDFont chineseFont = PDType0Font.load(document, fontStream); // 加载中文字体
                contentStream.setFont(chineseFont, fontSize);

                float y = initialY; // 起始Y坐标
                for (XWPFParagraph paragraph : wordDoc.getParagraphs()) {
                    String text = paragraph.getText();
                    if (!text.isEmpty()) {
                        contentStream.beginText();
                        contentStream.newLineAtOffset(marginLeft, y);
                        contentStream.showText(text);
                        contentStream.endText();
                        y -= (fontSize + 5); // 行间距
                    }
                    if (y < marginTop) { // 换页逻辑
                        contentStream.close(); // 关闭当前内容流
                        page = new PDPage();
                        document.addPage(page);
                        contentStream = new PDPageContentStream(document, page); // 创建新的内容流
                        y = initialY;
                    }
                }
            } finally {
                contentStream.close(); // 确保在结束时关闭内容流
            }
            document.save(outputPath);
            return new File(outputPath);
        } catch (IOException e) {
            throw new FileConverterException("Word 转 PDF 失败：" + e.getMessage());
        }
    }

    private File convertExcelToPdf(File inputFile, String outputPath) {
        try (PDDocument document = new PDDocument();
             Workbook workbook = WorkbookFactory.create(inputFile)) {
            Sheet sheet = workbook.getSheetAt(0);
            PDPage page = new PDPage();
            document.addPage(page);

            PDPageContentStream contentStream = new PDPageContentStream(document, page);
            try {
                // 设置字体
                InputStream fontStream = new FileInputStream("./resources/fonts/simhei.ttf");
                PDFont chineseFont = PDType0Font.load(document, fontStream); // 加载中文字体
                contentStream.setFont(chineseFont, fontSize);

                float y = initialY; // 起始Y坐标
                for (Row row : sheet) {
                    StringBuilder rowText = new StringBuilder();
                    for (Cell cell : row) {
                        rowText.append(cell.toString()).append("\t");
                    }
                    contentStream.beginText();
                    contentStream.newLineAtOffset(marginLeft, y);
                    contentStream.showText(rowText.toString());
                    contentStream.endText();
                    y -= (fontSize + 5); // 行间距

                    if (y < marginTop) { // 换页逻辑
                        contentStream.close();
                        page = new PDPage();
                        document.addPage(page);
                        contentStream = new PDPageContentStream(document, page);
                        y = initialY;
                    }
                }
            } finally {
                contentStream.close();
            }
            document.save(outputPath);
            return new File(outputPath);
        } catch (IOException e) {
            throw new FileConverterException("Excel 转 PDF 失败：" + e.getMessage());
        }
    }
}
