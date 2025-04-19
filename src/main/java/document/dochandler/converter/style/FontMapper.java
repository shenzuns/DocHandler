package document.dochandler.converter.style;

import org.apache.pdfbox.pdmodel.font.PDFont;

public class FontMapper {

    public static String mapFont(String pdfFontName) {
        return pdfFontName;
    }

    public static boolean isBold(PDFont pdfFont) {
        try {
            return pdfFont.getName().toLowerCase().contains("bold");
        }catch (Exception e) {
            return false;
        }
    }
    public static boolean isItalic(PDFont font) {
        try {
            return font.getName().toLowerCase().contains("italic")
                    || font.getName().toLowerCase().contains("oblique");
        }catch (Exception e) {
            return false;
        }
    }
}
