package net.summer.fastexcel.stream.exceptions;

public class OpenException extends RuntimeException {

    public OpenException(final String msg, final Exception e) {
        super(msg, e);
    }
}
