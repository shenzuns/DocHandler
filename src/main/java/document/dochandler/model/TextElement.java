package document.dochandler.model;

import java.awt.*;

public class TextElement extends PageElement{
    private String text;
    private String fontFamily;
    private float fontSize;
    private Color color;
    private boolean isBold;
    private boolean isItalic;
    private float pdfLeftMargin; // PDF内容区域左边距（单位：点）
    private float lineSpacing;
    public TextElement(String text, String fontFamily, float fontSize,
                       Color color, boolean isBold, boolean isItalic,
                       float x, float y, int pageNumber, float pageWidth,
                       float pdfLeftMargin, float lineSpacing) {
        this.text = text;
        this.fontFamily = fontFamily;
        this.fontSize = fontSize;
        this.color = color;
        this.isBold = isBold;
        this.isItalic = isItalic;
        this.x = x;
        this.y = y;
        this.pageNumber = pageNumber;
        this.pageWidth = pageWidth;
        this.pdfLeftMargin = pdfLeftMargin;
        this.lineSpacing = lineSpacing;
    }

    public String getText() {
        return text;
    }

    public String getFontFamily() {
        return fontFamily;
    }

    public float getFontSize() {
        return fontSize;
    }

    public Color getColor() {
        return color;
    }

    public boolean isBold() {
        return isBold;
    }

    public boolean isItalic() {
        return isItalic;
    }

    public float getPdfLeftMargin() {
        return pdfLeftMargin;
    }

    public float getLineSpacing() {
        return lineSpacing;
    }
}
