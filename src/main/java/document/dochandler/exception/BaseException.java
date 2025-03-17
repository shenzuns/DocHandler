package document.dochandler.exception;

public class BaseException extends RuntimeException {
    /**
     * 自定义异常基类
     */
    public BaseException(String message) {
        super(message);
    }
}
