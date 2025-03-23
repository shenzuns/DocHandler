package document.dochandler.converter.impl;

import document.dochandler.converter.FileConverter;
import document.dochandler.exception.BaseException;
import document.dochandler.exception.FileConverterException;
import document.dochandler.utils.FileValidatorUtils;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;


import java.io.*;
import java.util.List;
public class ToExcelConverter implements FileConverter {
    /**
     * 将文件转换为 Excel 类型
     *
     * @param inputFile  输入文件
     * @param outputPath 输出路径，若为空则生成在输入文件所在目录
     * @return 转换后的 Excel 文件
     * @throws FileConverterException 如果转换失败
     */
    @Override
    public File ToExcelConvert(File inputFile, String outputPath) {
        try {
            if (!FileValidatorUtils.isFileValid(inputFile)) {
                throw new Exception("输入文件无效");
            }
            if (outputPath == null || outputPath.isEmpty()) {
                outputPath = inputFile.getParent() + File.separator +
                        inputFile.getName().replaceFirst("[.][^.]+$", "") + ".xlsx";
            }

            if (inputFile.getName().endsWith(".txt")) {
                convertTxtToExcel(inputFile, outputPath);
            } else if (inputFile.getName().endsWith(".pdf")) {
                convertPdfToExcel(inputFile, outputPath);
            } else if (inputFile.getName().endsWith(".doc") || inputFile.getName().endsWith(".docx")) {
                convertWordToExcel(inputFile, outputPath);
            }

            // 返回目标文件
            return new File(outputPath);
        } catch (Exception e) {
            throw new FileConverterException("文件转换失败: " + e.getMessage());
        }
    }

    @Override
    public File ToWordConvert(File inputFile, String outputPath) {
        throw new BaseException("该实现类仅支持转excel");
    }
    @Override
    public File ToPdfConvert(File inputFile, String outputPath) {
        throw new BaseException("该实现类仅支持转excel");
    }
    @Override
    public File ToJsonConvert(File inputFile, String outputPath) {
        throw new BaseException("该实现类仅支持转excel");
    }

    private void convertTxtToExcel(File inputFile, String outputPath) {
        try (BufferedReader reader = new BufferedReader(new FileReader(inputFile));
             Workbook workbook = new XSSFWorkbook(); // 创建 Excel Workbook
             FileOutputStream fos = new FileOutputStream(outputPath)) {

            Sheet sheet = workbook.createSheet("Sheet1");
            String line;
            int rowIndex = 0;

            while ((line = reader.readLine()) != null) {
                // 判断是否可能是表格信息（比如用“，”或“\t”分隔的内容）
                if (line.contains(",") || line.contains("\t")) {
                    Row row = sheet.createRow(rowIndex++);
                    String[] columns = line.split("[,\t]");
                    for (int i = 0; i < columns.length; i++) {
                        row.createCell(i).setCellValue(columns[i].trim());
                    }
                }
            }

            workbook.write(fos);
        } catch (IOException e) {
            throw new FileConverterException("TXT 文件转换为 Excel 出错：" + e.getMessage());
        }
    }

    private void convertPdfToExcel(File inputFile, String outputPath) {
        try (PDDocument document = PDDocument.load(inputFile);
             Workbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(outputPath)) {

            Sheet sheet = workbook.createSheet("Sheet1");
            PDFTextStripper stripper = new PDFTextStripper();

            for (int pageNum = 0; pageNum < document.getNumberOfPages(); ++pageNum) {
                stripper.setStartPage(pageNum + 1);
                stripper.setEndPage(pageNum + 1);

                String text = stripper.getText(document);
                String[] lines = text.split("\\r?\\n");
                int rowIndex = sheet.getLastRowNum() + 1;

                for (String line : lines) {
                    // 只处理可能包含表格的行
                    if (isTableRow(line)) {
                        Row row = sheet.createRow(rowIndex++);
                        String[] columns = splitToColumns(line);
                        for (int i = 0; i < columns.length; i++) {
                            row.createCell(i).setCellValue(columns[i].trim());
                        }
                    }
                }
            }
            workbook.write(fos);
        } catch (IOException e) {
            throw new FileConverterException("PDF 提取表格转换为 Excel 失败：" + e.getMessage());
        }
    }

    private void convertWordToExcel(File inputFile, String outputPath) {
        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(outputPath)) {

            XWPFDocument document = new XWPFDocument(fis);
            Sheet sheet = workbook.createSheet("Sheet1");
            int rowIndex = 0;

            // 仅提取表格
            for (XWPFTable table : document.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    Row excelRow = sheet.createRow(rowIndex++);
                    List<XWPFTableCell> cells = row.getTableCells();
                    for (int i = 0; i < cells.size(); i++) {
                        excelRow.createCell(i).setCellValue(cells.get(i).getText());
                    }
                }
            }
            workbook.write(fos);
        } catch (IOException e) {
            throw new FileConverterException("Word 提取表格转换为 Excel 失败：" + e.getMessage());
        }
    }


    private String[] splitToColumns(String line) {
        return line.split("\\s{2,}"); // 按连续空格分割
    }

    private boolean isTableRow(String line) {
        return line.matches("^\\d+.*\\d+$"); // 示例：行首末都有数字
    }

}
