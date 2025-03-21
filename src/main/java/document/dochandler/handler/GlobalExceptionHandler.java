package document.dochandler.handler;
import document.dochandler.exception.BaseException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.RestControllerAdvice;
@RestControllerAdvice
public class GlobalExceptionHandler {
    private static final Logger logger = LoggerFactory.getLogger(GlobalExceptionHandler.class);

    @ExceptionHandler(BaseException.class)
    public void handleBaseException(BaseException ex) {
        // 记录日志
        logger.error("发生业务异常: {}", ex.getMessage(), ex);
    }

    @ExceptionHandler(Exception.class)
    public void fileConverterException(Exception ex) {
        logger.error("文件转换失败: {}", ex.getMessage(), ex);
    }

    @ExceptionHandler(Exception.class)
    public void FileTestException(Exception ex) {
        logger.error("文件转换失败: {}", ex.getMessage(), ex);
    }
}
