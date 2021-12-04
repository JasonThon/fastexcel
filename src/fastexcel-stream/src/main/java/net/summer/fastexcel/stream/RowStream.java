package net.summer.fastexcel.stream;

import java.io.Closeable;
import java.io.IOException;
import java.util.Iterator;
import java.util.function.Function;

/**
 * @author songyi
 * @Date 2021/11/30
 **/
public abstract class RowStream<T> implements Iterator<CellValue<T>>, MapStream<T>, Closeable {
    public abstract T[] asSlice();

    @Override
    public void close() throws IOException {

    }

    @Override
    public <R> MapStream<R> map(final Function<? super T, ? extends R> mapper) {
        return null;
    }

    public abstract <R> void row(R rawData);
}
