package document.dochandler.model;

public abstract class PageElement {
    protected float x;        // X坐标（单位：点）
    protected float y;        // Y坐标（点）
    protected float pageWidth;    // 宽度（点）
    protected float pageHeight;   // 高度（点）
    protected int pageNumber; // 页码

    public float getX() {
        return x;
    }

    public float getY() {
        return y;
    }
    public float getPageWidth() {
        return pageWidth;
    }
    public float getPageHeight() {
        return pageHeight;
    }
    public int getPageNumber() {
        return pageNumber;
    }
}
