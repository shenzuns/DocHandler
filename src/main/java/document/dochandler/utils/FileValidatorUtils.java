package document.dochandler.utils;

import java.io.File;

public class FileValidatorUtils {
    public static boolean isFileValid(File file) {
        if (file == null || !file.exists() ||!file.isFile()) {
            return false;
        }
        return true;
    }
}
