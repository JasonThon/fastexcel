package net.summer.fastexcel.stream.exceptions;

public class ReadException extends RuntimeException {

    public ReadException(final String msg, final Exception e) {
        super(msg, e);
    }
}
