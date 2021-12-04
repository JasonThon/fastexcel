package net.summer.fastexcel.stream;

import java.util.function.Consumer;

/**
 * @author songyi
 * @Date 2021/12/3
 **/
public abstract class Table2DStream<T> implements TableStream<RowStream<T>> {
    final Integer limit;
    private int counter;

    protected Table2DStream(final Integer limit) {
        this.limit = limit;
    }

    @Override
    public void foreach(Consumer<RowStream<T>> consumer) {
        if (limit != null && counter < limit) {
            TableStream.super.foreach(consumer);
            counter++;
        }
    }
}
