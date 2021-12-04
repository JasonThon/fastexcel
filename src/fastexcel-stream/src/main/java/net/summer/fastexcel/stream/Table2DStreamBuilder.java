package net.summer.fastexcel.stream;

/**
 * @author songyi
 * @Date 2021/12/1
 **/
public interface Table2DStreamBuilder<T> {
    <R extends Table2DStream<T>> R build();
}
