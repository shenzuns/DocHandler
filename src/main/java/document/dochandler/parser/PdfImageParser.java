package document.dochandler.parser;

import document.dochandler.model.ImageElement;
import org.apache.pdfbox.contentstream.PDFStreamEngine;
import org.apache.pdfbox.contentstream.operator.Operator;
import org.apache.pdfbox.cos.COSBase;
import org.apache.pdfbox.cos.COSName;
import org.apache.pdfbox.cos.COSNumber;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.util.Matrix;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Stack;

public class PdfImageParser {
    private int currentPageNumber = 0;
    public List<ImageElement> parsePdfImages(File inputFile) throws IOException {
        List<ImageElement> imageElements = new ArrayList<>();
        try(PDDocument document = PDDocument.load(inputFile)) {
            for (PDPage page : document.getPages()) {
                parsePageImages(page, imageElements);
                currentPageNumber++;
            }
        }
        return imageElements;
    }

    private void parsePageImages(PDPage page, List<ImageElement> imageElements) throws IOException {
        for (var xobject : page.getResources().getXObjectNames()) {
            var obj = page.getResources().getXObject(xobject);
            if (obj instanceof PDImageXObject) {
                PDImageXObject image = (PDImageXObject) obj;
                Matrix ctm = parseMatrixFromContentStream(page, xobject);

                // 获取页面尺寸
                PDRectangle mediaBox = page.getMediaBox();
                float pageWidthCm = mediaBox.getWidth();
                float pageHeightCm = mediaBox.getHeight();

                // 计算实际坐标
                float xCm = ctm.getTranslateX();
                float yCm = ctm.getScalingFactorY()+129.0f;

                float wordYCm = pageHeightCm - yCm - image.getHeight();
//                System.out.println("xCm + Ycm" + xCm + " " + yCm);
                ImageElement imageElement = new ImageElement(
                        getImageBytes(image),
                        xCm,
                        yCm,
                        ctm.getScaleX(),
                        ctm.getScaleY(),
                        currentPageNumber,
                        pageWidthCm,
                        pageHeightCm
                );
                imageElements.add(imageElement);
            }
        }
    }

    private Matrix parseMatrixFromContentStream(PDPage page, COSName xobjectName) throws IOException {
        PDFStreamEngine engine = new PDFStreamEngine() {
            private final Stack<Matrix> matrixStack = new Stack<>();
            private Matrix currentMatrix = new Matrix();
            private  COSName targetXObject = xobjectName;

            @Override
            protected void processOperator(Operator operator, List<COSBase> operands) throws IOException {
                // 跟踪变换矩阵操作
                switch (operator.getName()) {
                    case "q":
                        matrixStack.push(currentMatrix.clone());
                        break;
                    case "Q":
                        if (!matrixStack.isEmpty()) {
                            currentMatrix = matrixStack.pop();
                        }
                        break;
                    case "cm":
                        if (operands.size() >= 6) {
                            Matrix cmMatrix = new Matrix(
                                    ((COSNumber) operands.get(0)).floatValue(),  // a
                                    ((COSNumber) operands.get(1)).floatValue(),  // b
                                    ((COSNumber) operands.get(2)).floatValue(),  // c
                                    ((COSNumber) operands.get(3)).floatValue(),  // d
                                    ((COSNumber) operands.get(4)).floatValue(),  // e
                                    ((COSNumber) operands.get(5)).floatValue()   // f
                            );
//                            System.out.print("a: " + ((COSNumber) operands.get(0)).floatValue() +
//                                    " \nb: " + ((COSNumber) operands.get(1)).floatValue() +
//                                    " \nc: " + ((COSNumber) operands.get(2)).floatValue() +
//                                    " \nd: " + ((COSNumber) operands.get(3)).floatValue() +
//                                    " \ne: " + ((COSNumber) operands.get(4)).floatValue() +
//                                    " \nf: " + ((COSNumber) operands.get(5)).floatValue() + "----\n");
                            currentMatrix = currentMatrix.multiply(cmMatrix);
                        }
                        break;
                    case "Do":
                        COSName name = (COSName) operands.get(0);
                        if (name.equals(targetXObject)) {
                            throw new MatrixFoundException(currentMatrix);
                        }
                        break;
                }
                super.processOperator(operator, operands);
            }


        };

        try {
            engine.processPage(page);
            return new Matrix(); // 未找到时返回单位矩阵
        } catch (MatrixFoundException e) {
            return e.matrix;
        }
    }
    static class MatrixFoundException extends RuntimeException {
        final Matrix matrix;
        MatrixFoundException(Matrix matrix) {
            this.matrix = matrix;
        }
    }

    private byte[] getImageBytes(PDImageXObject image) throws IOException {
        BufferedImage bufferedImage = image.getImage();
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(bufferedImage, "png", baos);
        return baos.toByteArray();
    }
}
