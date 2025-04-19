package document.dochandler.writer;


import document.dochandler.model.ImageElement;
import document.dochandler.model.PageElement;
import document.dochandler.model.TextElement;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.LineSpacingRule;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.awt.*;
import java.io.ByteArrayInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class WordDocumentWriter {
    private static final float EPSILON = 1.0f;
    private XWPFParagraph currentParagraph; // 当前段落
    private float currentY;                 // 当前处理的Y坐标（PDF坐标系）
    private int currentPage;                // 当前处理的页码
    private XWPFRun currentRun;             // 当前处理的Run
    public void createDocument(List<TextElement> textElements,
                               List<ImageElement> imageElements,
                               String outputFilePath) throws Exception {
        //合并元素并排序
        List<PageElement> allElements = new ArrayList<>();
        allElements.addAll(textElements);
        allElements.addAll(imageElements);
        allElements.sort((a, b) -> {
            if (a.getPageNumber() != b.getPageNumber()) {
                return Integer.compare(a.getPageNumber(), b.getPageNumber());
            }
            return Float.compare(a.getY(), b.getY());
        });

        try (XWPFDocument document = new XWPFDocument()) {
            for (PageElement element : allElements) {
                if (element instanceof TextElement) {
                    processTextElement((TextElement) element, document);
                }else if (element instanceof ImageElement) {
                    processImageElement((ImageElement) element, document);
                }
            }
            // 保存文档
            try (FileOutputStream out = new FileOutputStream(outputFilePath)) {
                document.write(out);
            }
        }
    }

    private void processImageElement(ImageElement image, XWPFDocument document) {
        // 复用当前段落（如果Y坐标匹配）
            if (currentParagraph == null || Math.abs(image.getY() - currentY) > EPSILON) {
                currentParagraph = document.createParagraph();
                currentY = image.getY();
        }

        int indent = (int) (image.getX() * 20 - 1800);
        currentParagraph.setIndentationLeft(indent);
        // 插入图片前添加制表符
//        XWPFRun tabRun = currentParagraph.createRun();
//        tabRun.addTab();

        // 插入图片
        XWPFRun run = currentParagraph.createRun();
        try (ByteArrayInputStream bis = new ByteArrayInputStream(image.getImageData())) {
            run.addPicture(
                    bis,
                    XWPFDocument.PICTURE_TYPE_PNG,
                    "image.png",
                    Units.toEMU(image.getWidth()),
                    Units.toEMU(image.getHeight())
            );

        }catch (Exception e) {
            throw new RuntimeException("图片插入失败", e);
        }
        // 更新当前处理位置
        currentY = image.getY();

    }

    private void processTextElement(TextElement element, XWPFDocument document) {
//        System.out.print("textY: " + element.getText() + element.getY() + "\n");
        // 判断是否需要新段落
        boolean needNewPara = currentParagraph == null ||
                Math.abs(element.getY() - currentY) > EPSILON;

        if (needNewPara) {
            currentParagraph = document.createParagraph();
            //setRightTabStop(currentParagraph); // 设置右对齐制表位
            currentParagraph.setSpacingAfter(0);
            currentParagraph.setSpacingBefore(0);
            currentParagraph.setSpacingBetween(
                    1.5, // 行距倍数
                    LineSpacingRule.AUTO // 自动调整
            );
            //currentPage = element.getPageNumber();
//            currentParagraph.setAlignment(ParagraphAlignment.LEFT);
            // 计算缩进：PDF坐标转Word缩进（需考虑PDF边距和Word边距）
            float pdfX = element.getX();
            int indentTwips = (int) (pdfX * 20) - 1800; // 转换为缇并减去Word默认左边距
            currentParagraph.setIndentationLeft(indentTwips);

            currentY = element.getY();
            currentRun = null; // 重置当前Run
        }

        if (shouldCreateNewRun(currentRun, element)) {
            currentRun = currentParagraph.createRun();
            applyStyles(currentRun, element);
        }

        currentRun.setText(element.getText());
    }
    private boolean shouldCreateNewRun(XWPFRun currentRun, TextElement element) {
        return currentRun == null || !stylesMatch(currentRun, element);
    }
    private boolean stylesMatch(XWPFRun run, TextElement element) {
        String runColor = run.getColor();
        String elementColorHex = rgbToHex(element.getColor());
        return run.getFontFamily().equals(element.getFontFamily())
                && run.getFontSize() == element.getFontSize()
                && runColor.equals(elementColorHex);
    }
    private void applyStyles(XWPFRun run, TextElement element) {
        run.setFontFamily(element.getFontFamily());
        run.setFontSize(element.getFontSize());
        run.setColor(rgbToHex(element.getColor()));
        run.setBold(element.isBold());
        run.setItalic(element.isItalic());
    }
    private String rgbToHex(Color color) {
        return String.format("%02X%02X%02X", color.getRed(), color.getGreen(), color.getBlue());
    }
//    private void setRightTabStop(XWPFParagraph paragraph) {
//        CTP ctp = paragraph.getCTP();
//        CTPPr pPr = ctp.getPPr() != null ? ctp.getPPr() : ctp.addNewPPr();
//        CTTabs tabs = pPr.getTabs() != null ? pPr.getTabs() : pPr.addNewTabs();
//
//        // 假设页面总宽度为12240缇（8.5英寸），右边距1800缇
//        CTTabStop tabStop = tabs.addNewTab();
//        tabStop.setVal(STTabJc.RIGHT);
//        tabStop.setPos(BigInteger.valueOf(10440)); // 12240 - 1800 = 10440
//    }
}