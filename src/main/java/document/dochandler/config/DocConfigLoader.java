package document.dochandler.config;

import org.springframework.stereotype.Component;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;
@Component
public class DocConfigLoader {
    private Properties properties;

    public DocConfigLoader() {
        properties = new Properties();
        try (FileInputStream input = new FileInputStream("src/main/resources/application.properties")) {
            properties.load(input);
        } catch (IOException e) {
            throw new RuntimeException("Failed to load configuration file.", e);
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
    public float getPdfLineSpacing() {
        return Float.parseFloat(properties.getProperty("pdf.lineSpacing"));
    }
}
