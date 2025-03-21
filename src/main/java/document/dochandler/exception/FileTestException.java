package document.dochandler.exception;

public class FileTestException extends RuntimeException {
    private int errorCode;

    public FileTestException(String message) {
        super(message);
    }

    public FileTestException(String message, Throwable cause) {
        super(message, cause);
    }

    public FileTestException(int errorCode, String message) {
        super(message);
        this.errorCode = errorCode;
    }

    public int getErrorCode() {
        return errorCode;
    }
}
