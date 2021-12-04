package net.summer.fastexcel.stream.exceptions;

public class CloseException extends RuntimeException {

    public CloseException(final Exception e) {
        super(e);
    }

}
