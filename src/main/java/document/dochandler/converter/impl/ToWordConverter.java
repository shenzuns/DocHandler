package document.dochandler.converter.impl;

import document.dochandler.config.DocConfigLoader;
import document.dochandler.converter.FileConverter;
import document.dochandler.exception.BaseException;
import document.dochandler.exception.FileConverterException;
import document.dochandler.utils.FileValidatorUtils;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.io.*;

public class ToWordConverter implements FileConverter {
    private static DocConfigLoader docConfigLoader;

    public ToWordConverter(DocConfigLoader docConfigLoader) {
        this.docConfigLoader = docConfigLoader;
    }
    @Override
    public File ToExcelConvert(File inputFile, String outputPath) {
        throw new BaseException("该实现类仅支持转PDF");
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

    private void convertExcelToWord(File inputFile, String outputPath) throws IOException {
        FileInputStream excelFile = new FileInputStream(inputFile);
        Workbook workbook = WorkbookFactory.create(excelFile);

        XWPFDocument document = new XWPFDocument();

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
    }

    private void convertPdfToWord(File inputFile, String outputPath) throws IOException {
        try (PDDocument document = PDDocument.load(inputFile)) {
            PDFTextStripper stripper = new PDFTextStripper();
            String pdfText = stripper.getText(document);
            double lineSpacing = docConfigLoader.getWordLineSpacing();

            try (XWPFDocument wordDocument = new XWPFDocument()) {
                XWPFParagraph paragraph = wordDocument.createParagraph();
                XWPFRun run = paragraph.createRun();

                run.setText(pdfText);
                paragraph.setSpacingBetween(lineSpacing);

                try (FileOutputStream out = new FileOutputStream(outputPath)) {
                    wordDocument.write(out);
                }
            }
        } catch (Exception e) {
            throw new IOException("PDF 转换为 Word 失败：" + e.getMessage());
        }
    }

    private void convertTxtToWord(File inputFile, String outputPath) throws IOException {
        try (XWPFDocument document = new XWPFDocument()) {
            try (BufferedReader reader = new BufferedReader(new FileReader(inputFile))) {
                double lineSpacing = docConfigLoader.getWordLineSpacing();

                String line;
                while ((line = reader.readLine()) != null) {
                    XWPFParagraph paragraph = document.createParagraph();
                    XWPFRun run = paragraph.createRun();

                    run.setText(line);
                    paragraph.setSpacingBetween(lineSpacing);
                }
            }
            try (FileOutputStream out = new FileOutputStream(outputPath)) {
                document.write(out);
            }
        } catch (IOException e) {
            throw new IOException("TXT 转换为 Word 失败：" + e.getMessage());
        }
    }
}
