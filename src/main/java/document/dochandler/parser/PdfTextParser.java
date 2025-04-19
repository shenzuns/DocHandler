package document.dochandler.parser;

import document.dochandler.converter.style.FontMapper;
import document.dochandler.model.TextElement;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.TextPosition;

import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
public class PdfTextParser extends PDFTextStripper {
    private final List<TextElement> textElements = new ArrayList<>();
    private int currentPage = 0;
    private float currentPdfPageWidth;
    private float currentPdfLeftMargin;
    // 存储上一行的 y 坐标，用于计算行间距
    private float previousY = Float.NaN;
    public PdfTextParser() throws IOException {
        super();
        // 设置必要的参数（可选）
        this.setSortByPosition(true);    // 按位置排序文本
        this.setShouldSeparateByBeads(false);
        this.setLineSeparator("\n");      // 定义行分隔符
        this.setWordSeparator(" ");       // 定义单词分隔符
    }

    public List<TextElement> parsePdfText(File pdfFile) throws IOException {
        try (PDDocument document = PDDocument.load(pdfFile)) {
            this.setStartPage(1);
            this.setEndPage(document.getNumberOfPages());

            this.getText(document);
        }
        return textElements;
    }
    @Override
    protected void startPage(PDPage page) {
        // 更新当前页码（转换为 0-based）
        currentPage = getCurrentPageNo() - 1;
        // 获取页面尺寸和边距
        PDRectangle cropBox = page.getCropBox();
        float pageWidth = cropBox.getWidth();    // 页面内容区域宽度（点）
        float pageHeight = cropBox.getHeight();   // 页面内容区域高度（点）
        float pdfLeftMargin = cropBox.getLowerLeftX(); // 左边距（点）

        // 将页面信息传递给TextElement
        currentPdfPageWidth = pageWidth;
        currentPdfLeftMargin = pdfLeftMargin;
        previousY = Float.NaN;
    }
    @Override
    protected void writeString(String text, List<TextPosition> textPositions) {
        for (TextPosition textPosition : textPositions) {
            float currentY = textPosition.getYDirAdj();
            // 计算行间距
            float lineSpacing = Float.NaN;
            if (!Float.isNaN(previousY)){
                lineSpacing = currentY - previousY;
            }
            previousY = currentY;

            TextElement textElement = new TextElement(
                    textPosition.getUnicode(),
                    FontMapper.mapFont(textPosition.getFont().getName()),
                    textPosition.getFontSize(),
                    Color.BLACK,
                    FontMapper.isBold(textPosition.getFont()),
                    FontMapper.isItalic(textPosition.getFont()),
                    textPosition.getXDirAdj(),
                    textPosition.getYDirAdj(),
                    currentPage,
                    currentPdfPageWidth,
                    currentPdfLeftMargin,
                    lineSpacing
            );
            textElements.add(textElement);
        }
    }

    private Color extractColor(TextPosition tp) {

        return Color.BLACK;
    }

}