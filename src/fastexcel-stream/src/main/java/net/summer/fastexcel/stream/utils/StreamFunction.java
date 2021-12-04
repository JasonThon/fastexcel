package net.summer.fastexcel.stream.utils;

/**
 * @author songyi
 * @Date 2021/12/4
 **/
interface StreamFunction<K, V> {
    V apply(final K key, final int index);
}
