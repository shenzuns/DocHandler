package document.dochandler.converter.impl;

import document.dochandler.config.DocConfigLoader;
import document.dochandler.converter.FileConverter;
import document.dochandler.exception.BaseException;
import document.dochandler.exception.FileConverterException;
import document.dochandler.utils.FileValidatorUtils;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;

public class ToPDFConverter implements FileConverter {
    private DocConfigLoader configLoader;

    public ToPDFConverter(DocConfigLoader configLoader) {
        this.configLoader = configLoader;
    }
    String fontName = configLoader.get("pdf.font");
    float fontSize = configLoader.getInt("pdf.fontSize");
    float marginTop = configLoader.getInt("pdf.marginTop");
    float marginLeft = configLoader.getInt("pdf.marginLeft");
    float initialY = 800 - marginTop;
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
        try (BufferedReader reader = new BufferedReader(new FileReader(inputFile))) {

            PDPage page = new PDPage();
            document.addPage(page);
            PDPageContentStream contentStream = new PDPageContentStream(document, page);
            try {
                contentStream.setFont(PDType1Font.COURIER, fontSize);
                float y = initialY; // 起始Y坐标

                String line;
                while ((line = reader.readLine()) != null) {
                    contentStream.beginText();
                    contentStream.newLineAtOffset(marginLeft, y);
                    contentStream.showText(line);
                    contentStream.endText();
                    y -= (fontSize + 3); // 行间距

                    if (y < marginLeft) { // 换页逻辑
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
            throw new FileConverterException("TXT 转 PDF 失败：" + e.getMessage());
        }
    }

    private File convertWordToPdf(File inputFile, String outputPath) {
        try (PDDocument document = new PDDocument();
             FileInputStream inputStream = new FileInputStream(inputFile)) {
            XWPFDocument wordDoc = new XWPFDocument(inputStream);
            PDPage page = new PDPage();
            document.addPage(page);
            PDPageContentStream contentStream = new PDPageContentStream(document, page);

            try {
                contentStream.setFont(PDType1Font.COURIER, fontSize);
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
                contentStream.setFont(PDType1Font.COURIER, fontSize);
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
