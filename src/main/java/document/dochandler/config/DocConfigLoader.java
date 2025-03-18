package document.dochandler.config;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

public class DocConfigLoader {
    private Properties properties;

    public DocConfigLoader(String configFilePath) {
        properties = new Properties();
        try (FileInputStream input = new FileInputStream(configFilePath)) {
            properties.load(input);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public String get(String key) {
        return properties.getProperty(key);
    }

    public int getInt(String key) {
        return Integer.parseInt(properties.getProperty(key));
    }

    public double getDouble(String key) {
        return Double.parseDouble(properties.getProperty(key));
    }

    public double getWordLineSpacing() {
        return getDouble("word.lineSpacing");
    }

    public String getPdfFont() {
        return get("pdf.font");
    }

    public int getPdfFontSize() {
        return getInt("pdf.fontSize");
    }

    public int getPdfMarginTop() {
        return getInt("pdf.marginTop");
    }

    public int getPdfMarginBottom() {
        return getInt("pdf.marginBottom");
    }

    public int getPdfMarginLeft() {
        return getInt("pdf.marginLeft");
    }

    public int getPdfMarginRight() {
        return getInt("pdf.marginRight");
    }
}
