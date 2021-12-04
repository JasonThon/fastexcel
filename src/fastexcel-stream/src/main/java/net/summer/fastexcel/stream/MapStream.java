package net.summer.fastexcel.stream;

import java.util.function.Function;

/**
 * @author songyi
 * @Date 2021/11/30
 **/
public interface MapStream<T> {
    <R> MapStream<R> map(Function<? super T, ? extends R> mapper);

}
