package document.dochandler;

import document.dochandler.converter.ToJsonConverter;
import org.junit.jupiter.api.Test;

import java.io.File;

public class ToJsonTest {
    @Test
    public void toJsonTest() {

        ToJsonConverter converter = new ToJsonConverter();
        File inputFile = new File("./doc/智控院反诈教育安排.xlsx");
        String outputPath = "./doc/智控院反诈教育安排.json";

        File jsonFile = converter.ToJsonConvert(inputFile, outputPath);

    }
}
