package document.dochandler.converter.impl;

import document.dochandler.config.DocConfigLoader;
import document.dochandler.converter.FileConverter;
import document.dochandler.exception.BaseException;
import document.dochandler.exception.FileConverterException;
import document.dochandler.utils.FileValidatorUtils;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import java.io.*;
@Component
public class ToWordConverter implements FileConverter {
    private final DocConfigLoader configLoader;
    private final String font;
    private final float fontSize;
    private final float marginTop;
    private final float marginLeft;
    private final float initialY;
    private final float lineSpacing;

    @Autowired
    public ToWordConverter(DocConfigLoader configLoader) {
        this.configLoader = configLoader;
        this.font = configLoader.getPdfFont();
        this.fontSize = configLoader.getPdfFontSize();
        this.marginTop = configLoader.getPdfMarginTop();
        this.marginLeft = configLoader.getPdfMarginLeft();
        this.initialY = 800 - marginTop;
        this.lineSpacing = configLoader.getPdfLineSpacing();
    }
     /**
     * 将文件转换为 Word 类型
     *
     * @param inputFile      输入文件
     * @param outputPath     输出文件路径
     * @return 转换后的文件对象
     * @throws FileConverterException 如果转换失败
     */
    @Override
    public File ToWordConvert(File inputFile, String outputPath) {
        try {
            if (!FileValidatorUtils.isFileValid(inputFile)) {
                throw new FileConverterException("输入文件无效");
            }

            if (outputPath == null || outputPath.isEmpty()) {
                outputPath = inputFile.getParent() + File.separator +
                        inputFile.getName().replaceFirst("[.][^.]+$", "") + ".docx";
            }

            if (inputFile.getName().endsWith(".txt")) {
                convertTxtToWord(inputFile, outputPath);
            } else if (inputFile.getName().endsWith(".pdf")) {
                convertPdfToWord(inputFile, outputPath);
            } else if (inputFile.getName().endsWith(".xlsx") || inputFile.getName().endsWith(".xls")) {
                convertExcelToWord(inputFile, outputPath);
            } else {
                throw new IllegalArgumentException("不支持的文件类型");
            }

            return new File(outputPath); // 返回生成的文件对象
        } catch (Exception e) {
            throw new FileConverterException("文件转换为Word失败：" + e.getMessage());
        }
    }

    @Override
    public File ToPdfConvert(File inputFile, String outputPath) {
        throw new BaseException("该实现类仅支持转PDF");
    }

    @Override
    public File ToJsonConvert(File inputFile, String outputPath) {
        throw new BaseException("该实现类仅支持转PDF");
    }
    @Override
    public File ToExcelConvert(File inputFile, String outputPath) {
        throw new BaseException("该实现类仅支持转PDF");
    }

    private void convertExcelToWord(File inputFile, String outputPath) {
        try {
            FileInputStream excelFile = new FileInputStream(inputFile);
            Workbook workbook = WorkbookFactory.create(excelFile);

            XWPFDocument document = new XWPFDocument();

            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setFontFamily(font);
            run.setFontSize(fontSize);
            paragraph.setSpacingBetween(lineSpacing);

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                XWPFTable table = document.createTable(sheet.getPhysicalNumberOfRows(), sheet.getRow(0).getPhysicalNumberOfCells());

                for (int rowIndex = 0; rowIndex < sheet.getPhysicalNumberOfRows(); rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    for (int colIndex = 0; colIndex < row.getPhysicalNumberOfCells(); colIndex++) {
                        Cell cell = row.getCell(colIndex);
                        String cellValue = cell.toString();

                        table.getRow(rowIndex).getCell(colIndex).setText(cellValue);
                    }
                }
                document.createParagraph();
            }

            try (FileOutputStream outputStream = new FileOutputStream(outputPath)) {
                document.write(outputStream);
            }

            workbook.close();
            excelFile.close();
            document.close();
        }catch (Exception e) {
            throw new FileConverterException("Excel 转换为 Word 失败：" + e.getMessage());
        }
    }

    private void convertPdfToWord(File inputFile, String outputPath) {
        try (PDDocument pdfDocument = PDDocument.load(inputFile);
             XWPFDocument wordDocument = new XWPFDocument()) {

            PDFTextStripperByArea stripper = new PDFTextStripperByArea();
            stripper.setSortByPosition(true); // 按位置排序字符
            PDFTextStripper textStripper = new PDFTextStripper();

            textStripper.setSortByPosition(true); // 按位置信息提取每个字符
            String pdfText = textStripper.getText(pdfDocument).trim();

            // 提取每一页的内容
            for (int pageIndex = 1; pageIndex <= pdfDocument.getNumberOfPages(); pageIndex++) {
                textStripper.setStartPage(pageIndex);
                textStripper.setEndPage(pageIndex);

                // 按字符解析这一页
                String pageText = textStripper.getText(pdfDocument);

                XWPFParagraph paragraph = wordDocument.createParagraph();
                XWPFRun run = paragraph.createRun();

                // 逐段、逐字进行格式控制
                for (char c : pageText.toCharArray()) {
                    if (isBold(c)) {
                        run.setBold(true); // 设置加粗
                    }
                    run.setFontFamily(font); // 设置字体
                    run.setFontSize(fontSize); // 设置默认字体大小
                    run.setText(String.valueOf(c)); // 添加字符

                    // 自行完成换行逻辑（写成新段落）的处理
                    if (c == '\n') {
                        paragraph = wordDocument.createParagraph();
                        run = paragraph.createRun();
                    }
                }
            }

            // 保存为 Word 文件
            try (FileOutputStream out = new FileOutputStream(outputPath)) {
                wordDocument.write(out);
            }
        } catch (Exception e) {
            throw new FileConverterException("PDF 转 Word 失败：" + e.getMessage());
        }
    }
    private boolean isBold(char c) {
        // 模拟判断加粗：根据字体或样式特性自行实现判断
        // 例如：特定区域、大文字、标识性字体等
        return Character.isUpperCase(c); // 示例：大写字母默认加粗
    }

    private void convertTxtToWord(File inputFile, String outputPath) {
        try (XWPFDocument document = new XWPFDocument()) {
            try (BufferedReader reader = new BufferedReader(new FileReader(inputFile))) {

                String line;
                while ((line = reader.readLine()) != null) {
                    XWPFParagraph paragraph = document.createParagraph();
                    XWPFRun run = paragraph.createRun();
                    run.setFontFamily(font);
                    run.setFontSize(fontSize);
                    run.setText(line);
//                    paragraph.setSpacingBetween(lineSpacing);
                }
            }
            try (FileOutputStream out = new FileOutputStream(outputPath)) {
                document.write(out);
            }
        } catch (IOException e) {
            throw new FileConverterException("TXT 转换为 Word 失败：" + e.getMessage());
        }
    }
}
