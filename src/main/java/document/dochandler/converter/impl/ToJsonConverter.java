package document.dochandler.converter.impl;

import com.fasterxml.jackson.databind.ObjectMapper;
import document.dochandler.converter.FileConverter;
import document.dochandler.exception.BaseException;
import document.dochandler.exception.FileConverterException;
import document.dochandler.utils.FileValidatorUtils;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

public class ToJsonConverter implements FileConverter {
    @Override
    public File ToExcelConvert(File inputFile, String outputPath) {
        throw new BaseException("该实现类仅支持转json文件");
    }

    @Override
    public File ToWordConvert(File inputFile, String outputPath) {
        throw new BaseException("该实现类仅支持转json文件");
    }

    @Override
    public File ToPdfConvert(File inputFile, String outputPath) {
        throw new BaseException("该实现类仅支持转json文件");
    }
    /**
     * 将文件转换为 json 数据
     *
     * @param inputFile      输入文件
     * @param outputPath     输出路径
     * @throws Exception 如果转换失败
     */
    @Override
    public File ToJsonConvert(File inputFile, String outputPath) {
        if (!FileValidatorUtils.isFileValid(inputFile)) {
            throw new FileConverterException("输入文件不存在！");
        }

        if (outputPath == null || outputPath.isEmpty()) {
            outputPath = inputFile.getParent() + File.separator +
                    inputFile.getName().replaceFirst("[.][^.]+$", "") + ".json";
        }

        List<Map<String, Object>> content = new ArrayList<>();

        if (inputFile.getName().endsWith(".doc") || inputFile.getName().endsWith(".docx")) {
            // Word 处理
            content = extractWordContent(inputFile);
        } else if (inputFile.getName().endsWith(".pdf")) {
            // PDF 处理
            content = extractPdfContent(inputFile);
        } else if (inputFile.getName().endsWith(".xls") || inputFile.getName().endsWith(".xlsx")) {
            // Excel 处理
            content = extractExcelContent(inputFile);
        } else {
            throw new FileConverterException("不支持的文件类型！");
        }

        File jsonFile = new File(outputPath);
        ObjectMapper mapper = new ObjectMapper();
        try {
            mapper.writeValue(jsonFile, content);
        }catch (IOException e) {
            throw new FileConverterException("转换为Json失败！");
        }
        return jsonFile;
    }
    private List<Map<String, Object>> extractExcelContent(File inputFile) {
        List<Map<String, Object>> content = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook workbook = new XSSFWorkbook(fis)) {

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);

                for (Row row : sheet) {
                    Map<String, Object> rowData = new HashMap<>();
                    List<String> cells = new ArrayList<>();
                    for (Cell cell : row) {
                        cell.setCellType(CellType.STRING);
                        cells.add(cell.getStringCellValue());

                    }
                    rowData.put("type", "row");
                    rowData.put("cells", cells);
                    content.add(rowData);
                }
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return content;
    }

    private List<Map<String, Object>> extractPdfContent(File inputFile) {
        List<Map<String, Object>> content = new ArrayList<>();

        try (PDDocument doc = PDDocument.load(inputFile)) {
            PDFTextStripper stripper = new PDFTextStripper();
            String text = stripper.getText(doc);
            String[] lines = text.split("\\r?\\n");
            for (String line : lines) {
                if (!line.isEmpty()) {
                    Map<String, Object> entry = new HashMap<>();
                    entry.put("text", "line");
                    entry.put("content", line);
                    content.add(entry);
                }
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        return content;
    }

    private List<Map<String, Object>> extractWordContent(File inputFile) {
        List<Map<String, Object>> content = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(inputFile);
             XWPFDocument doc = new XWPFDocument(fis)) {

            for (XWPFParagraph paragraph : doc.getParagraphs()) {
                if (!paragraph.getText().isEmpty()) {
                    Map<String, Object> entry = new HashMap<>();
                    entry.put("text", "paragraph");
                    entry.put("content", paragraph.getText());
                    content.add(entry);
                }
            }

            for (XWPFTable table : doc.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    Map<String, Object> rowData = new HashMap<>();
                    List<String> cells = row.getTableCells().stream()
                            .map(XWPFTableCell::getText)
                            .collect(Collectors.toList());
                    rowData.put("type", "tableRow");
                    rowData.put("cells", cells);
                    content.add(rowData);
                }
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        return content;
    }
}
