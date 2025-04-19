package document.dochandler.converter;


import document.dochandler.model.ImageElement;
import document.dochandler.model.TextElement;
import document.dochandler.parser.PdfImageParser;
import document.dochandler.parser.PdfTextParser;
import document.dochandler.writer.WordDocumentWriter;

import java.io.File;
import java.util.List;

public class PdfToWordConverter {
    public void convert(File pdfFile, String outputFilePath) throws Exception {
        PdfTextParser parser = new PdfTextParser();
        List<TextElement> elements = parser.parsePdfText(pdfFile);

        PdfImageParser imageParser = new PdfImageParser();
        List<ImageElement> images = imageParser.parsePdfImages(pdfFile);

        WordDocumentWriter writer = new WordDocumentWriter();
        writer.createDocument(elements, images, outputFilePath);
    }
}
