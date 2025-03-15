package document.dochandler.converter;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;

public class ToPDFConverter {
    /**
     * 将文件转换为 Word 类型
     *
     * @param inputFile      输入文件
     * @param outputPath     输出文件
     * @throws Exception 如果转换失败
     */
    public void toPdfHandler(File inputFile, String outputPath) throws Exception {
        if (inputFile == null || !inputFile.exists()) {
            throw new Exception("输入文件错误或不存在！");
        }
        if (outputPath == null || outputPath.isEmpty()) {
            outputPath = inputFile.getParent() + File.separator +
                    inputFile.getName().replaceFirst("[.][^.]+$", "") + ".pdf";
        }

        if (inputFile.getName().endsWith(".txt")) {
            convertTxtToPdf(inputFile, outputPath);
        } else if (inputFile.getName().endsWith(".xlsx")) {
            convertExcelToPdf(inputFile, outputPath);
        } else if (inputFile.getName().endsWith(".doc") || inputFile.getName().endsWith(".docx")) {
            convertWordToPdf(inputFile, outputPath);
        }
    }
    private void convertTxtToPdf (File inputFile, String outputPath) throws Exception {
        PDDocument document = new PDDocument();
        try (BufferedReader reader = new BufferedReader(new FileReader(inputFile))) {
            PDPage page = new PDPage();
            document.addPage(page);
            PDPageContentStream  contentStream = new PDPageContentStream(document, page);
            try {
                contentStream.setFont(PDType1Font.COURIER, 12);
                float y = 700; // 起始Y坐标

                String line;
                while ((line = reader.readLine()) != null) {
                    contentStream.beginText();
                    contentStream.newLineAtOffset(50, y);
                    contentStream.showText(line);
                    contentStream.endText();
                    y -= 15; // 行间距

                    if (y < 50) { // 换页逻辑
                        contentStream.close();
                        page = new PDPage();
                        document.addPage(page);
                        contentStream = new PDPageContentStream(document, page);
                        y = 700;
                    }
                }
            }finally {
                contentStream.close();
            }
            document.save(outputPath);
        } catch (IOException e) {
            throw new Exception("TXT转PDF失败: " + e.getMessage());
        }
    }
    private void convertWordToPdf(File inputFile, String outputPath) throws Exception {
        try (PDDocument document = new PDDocument();
             FileInputStream inputStream = new FileInputStream(inputFile)) {
            XWPFDocument wordDoc = new XWPFDocument(inputStream);
            PDPage page = new PDPage();
            document.addPage(page);
            PDPageContentStream contentStream = new PDPageContentStream(document, page);

            try {
                contentStream.setFont(PDType1Font.COURIER, 12);
                float y = 700; // 起始Y坐标
                for (XWPFParagraph paragraph : wordDoc.getParagraphs()) {
                    String text = paragraph.getText();
                    if (!text.isEmpty()) {
                        contentStream.beginText();
                        contentStream.newLineAtOffset(50, y);
                        contentStream.showText(paragraph.getText());
                        contentStream.endText();
                        y -= 15; // 行间距
                    }
                    if (y < 50) { // 换页逻辑
                        contentStream.close();
                        page = new PDPage();
                        document.addPage(page);
                        contentStream = new PDPageContentStream(document, page);
                        y = 700;
                    }
                }

                for (XWPFTable table : wordDoc.getTables()) {
                    for (XWPFTableRow tableRow : table.getRows()) {
                        StringBuilder rowText = new StringBuilder();
                        for (XWPFTableCell cell : tableRow.getTableCells()) {
                            rowText.append(cell.getText()).append(" | ");
                        }
                        contentStream.beginText();
                        contentStream.newLineAtOffset(50, y);
                        contentStream.showText(rowText.toString());
                        contentStream.endText();
                        y -= 15; // 行间距
                    }
                }

            } catch (IOException e) {
                throw new Exception("Word转PDF失败: " + e.getMessage());
            }
            document.save(outputPath);
        }
    }

    private void convertExcelToPdf(File inputFile, String outputPath) throws Exception {
        try (PDDocument document = new PDDocument();
             Workbook workbook = WorkbookFactory.create(inputFile)) {
            Sheet sheet = workbook.getSheetAt(0);
            PDPage page = new PDPage();
            document.addPage(page);

            PDPageContentStream contentStream = new PDPageContentStream(document, page);
            try {
                contentStream.setFont(PDType1Font.COURIER, 12);
                float y = 700; // 起始Y坐标
                for (Row row : sheet) {
                    StringBuilder rowText = new StringBuilder();
                    for (Cell cell : row) {
                        rowText.append(cell.toString()).append("\t");
                    }
                    contentStream.beginText();
                    contentStream.newLineAtOffset(50, y);
                    contentStream.showText(rowText.toString());
                    contentStream.endText();
                    y -= 15; // 行间距

                    if (y < 50) { // 换页逻辑
                        contentStream.close();
                        page = new PDPage();
                        document.addPage(page);
                        contentStream = new PDPageContentStream(document, page);
                        y = 700;
                    }
                }
                document.save(outputPath);
            } catch (IOException e) {
                throw new Exception("Excel转PDF失败: " + e.getMessage());
            }
        }

    }
}
