package document.dochandler.converter;

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

public class ToExcelConverter {
    /**
     * 将文件转换为 Excel 类型
     *
     * @param inputFile      输入文件
     * @param outputPath     输出路径
     * @throws Exception 如果转换失败
     */
    public void toExcelHandler(File inputFile, String outputPath)throws Exception {
        if (inputFile == null ||!inputFile.exists()) {
            throw new Exception("输入文件错误或不存在！");
        }
        if (outputPath == null || outputPath.isEmpty()) {
            outputPath = inputFile.getParent() + File.separator +
                    inputFile.getName().replaceFirst("[.][^.]+$", "") + ".xlsx";
        }

        if (inputFile.getName().endsWith(".txt")) {
            convertTxtToExcel(inputFile, outputPath);
        }else if (inputFile.getName().endsWith(".pdf")) {
            convertPdfToExcel(inputFile, outputPath);
        }else if (inputFile.getName().endsWith(".doc") || inputFile.getName().endsWith(".docx")) {
            convertWordToExcel(inputFile, outputPath);
        }
    }
    private void convertTxtToExcel(File inputFile, String outputPath) throws Exception {
        // 打开输入 TXT 文件
        try (BufferedReader reader = new BufferedReader(new FileReader(inputFile));
             Workbook workbook = new XSSFWorkbook(); // 创建 Excel Workbook
             FileOutputStream fos = new FileOutputStream(outputPath)) {

            // 创建 Excel 表(Sheet)
            Sheet sheet = workbook.createSheet("Sheet1");

            String line;
            int rowIndex = 0;

            // 按行读取 TXT 文件，将每行写入 Excel 的一行
            while ((line = reader.readLine()) != null) {
                Row row = sheet.createRow(rowIndex++);
                row.createCell(0).setCellValue(line);
            }

            // 将数据写入目标路径的 Excel 文件
            workbook.write(fos);
        } catch (IOException e) {
            throw new Exception("TXT 文件转换为 Excel 出错：" + e.getMessage());
        }
    }
    private void convertPdfToExcel(File inputFile, String outputPath) throws Exception {
        try (PDDocument document = PDDocument.load(inputFile);
             Workbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(outputPath)) {

            Sheet sheet = workbook.createSheet("Sheet1");
            PDFTextStripper stripper = new PDFTextStripper();

            // 按页处理PDF
            for (int pageNum = 0; pageNum < document.getNumberOfPages(); ++pageNum) {
                stripper.setStartPage(pageNum + 1);
                stripper.setEndPage(pageNum + 1);

                // 获取页面文本并按行分割
                String text = stripper.getText(document);
                String[] lines = text.split("\\r?\\n");

                int rowIndex = sheet.getLastRowNum() + 1;

                // 高级表格识别逻辑（示例）
                for (String line : lines) {
                    if (isTableRow(line)) { // 自定义表格行判断
                        Row row = sheet.createRow(rowIndex++);

                        // 按列分割逻辑（示例）
                        String[] columns = splitToColumns(line);
                        for (int i = 0; i < columns.length; i++) {
                            row.createCell(i).setCellValue(columns[i].trim());
                        }
                    }
                }
            }
            workbook.write(fos);
        } catch (IOException e) {
            throw new Exception("PDF转换Excel失败: " + e.getMessage());
        }
    }
    private void convertWordToExcel(File inputFile, String outputPath) throws Exception {
        try (FileInputStream fis = new FileInputStream(inputFile);
            Workbook workbook = WorkbookFactory.create(fis);
            FileOutputStream fos = new FileOutputStream(outputPath)) {

            XWPFDocument document = new XWPFDocument(fis);
            Sheet sheet = workbook.createSheet("Sheet1");
            int rowIndex = 0;

            for (XWPFParagraph paragraph : document.getParagraphs()) {
                String text = paragraph.getText();
                if (!text.isEmpty()) {
                    Row row = sheet.createRow(rowIndex++);
                    row.createCell(0).setCellValue(text);
                }
            }
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
        }catch (IOException e) {
            throw new Exception("Word转换Excel失败: " + e.getMessage());
        }
    }
    private String[] splitToColumns(String line) {
        return line.split("\\s{2,}"); // 按连续空格分割
    }
    // 表格行识别策略（示例）
    private boolean isTableRow(String line) {
        return line.matches("^\\d+.*\\d+$"); // 示例：行首末都有数字
    }

}
