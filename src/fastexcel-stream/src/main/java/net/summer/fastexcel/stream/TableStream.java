package net.summer.fastexcel.stream;

import java.util.Iterator;
import java.util.List;
import java.util.Objects;
import java.util.function.Consumer;

/**
 * @author songyi
 * @Date 2021/11/30
 **/
public interface TableStream<T> extends Iterator<T>, AutoCloseable {

    @Override
    default void close() {

    }

    default void foreach(Consumer<T> consumer) {
        while (hasNext()) {
            T next = next();
            Objects.requireNonNull(next);

            consumer.accept(next);
        }
    }

    List<String> tableNames();
}
