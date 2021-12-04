package net.summer.fastexcel.stream.utils;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.List;
import java.util.Objects;
import java.util.function.Function;
import java.util.function.Supplier;
import java.util.stream.Stream;

import com.google.common.collect.Lists;

/**
 * @author songyi
 * @Date 2021/12/4
 **/
public class CollectionUtils {
    public static <K, V> List<V> doIndexMap(final Collection<K> list,
                                            final StreamFunction<K, V> mapper,
                                            final boolean skipNull) {
        if (list == null) {
            return Collections.emptyList();
        }
        List<V> results = Lists.newArrayListWithExpectedSize(list.size());

        int index = 0;
        for (K k : list) {
            V apply = mapper.apply(k, index);
            index++;

            if (skipNull && apply == null) {
                continue;
            }

            results.add(apply);
        }

        return results;
    }

    public static <K, V> List<V> map(final List<K> list, final Function<K, V> mapper) {
        if (list == null) {
            return Lists.newArrayList();
        }

        return doIndexMap(list, (value, index) -> mapper.apply(value), false);
    }

    public static <T> T safeGet(final List<T> formats, final int index) {
        if (Objects.isNull(formats) || formats.size() <= index) {
            return null;
        }

        return formats.get(index);
    }

    public static <T> int size(final List<T> list) {
        if (list != null) {
            return list.size();
        }

        return 0;
    }

    public static <T> Stream<T> stream(final Collection<T> list) {
        return Objects.requireNonNullElseGet(list, (Supplier<Collection<T>>) ArrayList::new).stream();

    }
}
