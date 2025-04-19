package document.dochandler.model;

public class ImageElement extends PageElement{
    private byte[] imageData; //图片的二进制数据
    private float width;
    private float height;

    public ImageElement(byte[] imageData, float x, float y, float width,
                        float height, int pageNumber, float pageWidth, float pageHeight) {
        this.imageData = imageData;
        this.x = x;
        this.y = y;
        this.width = width;
        this.height = height;
        this.pageNumber = pageNumber;
        this.pageWidth = pageWidth;
        this.pageHeight = pageHeight;
    }

    public byte[] getImageData() {
        return imageData;
    }
    public float getWidth() {
        return width;
    }
    public float getHeight() {
        return height;
    }
}
