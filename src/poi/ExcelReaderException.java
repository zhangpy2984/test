package poi;

/**
 * @author zhangpengyue
 * @date 2018/12/16
 */
public class ExcelReaderException  extends Exception{

    public ExcelReaderException() {
    }

    public ExcelReaderException(String message) {
        super(message);
    }

    public ExcelReaderException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelReaderException(Throwable cause) {
        super(cause);
    }

    public ExcelReaderException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
