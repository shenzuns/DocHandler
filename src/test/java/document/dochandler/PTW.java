package document.dochandler;


import document.dochandler.converter.PdfToWordConverter;
import document.dochandler.model.ImageElement;
import document.dochandler.model.TextElement;
import document.dochandler.parser.PdfImageParser;
import document.dochandler.parser.PdfTextParser;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.util.List;

public class PTW {
    public static void main(String[] args) {

    }
    private static void pdf2word(String pathName, String outPath) {
        try {
            PdfToWordConverter converter = new PdfToWordConverter();
            converter.convert(
                    new File(pathName),
                    outPath
            );
            System.out.println("Conversion completed successfully!");
        } catch (Exception e) {
            System.err.println("Conversion failed: " + e.getMessage());
            e.printStackTrace();
        }
    }
    @Test
    public void test12() throws IOException {
        pdf2word("./doc/12.pdf", "./doc/output.docx");
    }
    @Test
    public void test11() throws IOException {
        pdf2word("./doc/11.pdf", "./doc/output.docx");
    }
    @Test
    public void testChu() throws IOException {
        pdf2word("./doc/出师表.pdf", "./doc/chu.docx");
    }
    @Test
    public void chu() throws Exception {
        PdfTextParser parser = new PdfTextParser();
        List<TextElement> elements = parser.parsePdfText(new File("./doc/出师表.pdf"));

        PdfImageParser imageParser = new PdfImageParser();
        List<ImageElement> images = imageParser.parsePdfImages(new File("./doc/出师表.pdf"));


    }
}
