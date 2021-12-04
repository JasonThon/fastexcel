package net.summer.fastexcel.stream.exceptions;

/**
 * @author songyi
 * @Date 2021/12/3
 **/
public class ExcelCommonException extends RuntimeException {
    public ExcelCommonException(final String message, final Throwable cause) {
        super(message, cause);
    }

    public ExcelCommonException(final String message) {
        super(message);
    }
}
