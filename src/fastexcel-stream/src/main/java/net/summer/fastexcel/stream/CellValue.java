package net.summer.fastexcel.stream;

import lombok.Getter;
import lombok.Setter;

/**
 * @author songyi
 * @Date 2021/11/30
 **/
public abstract class CellValue<T> {
    @Setter
    @Getter
    private T value;

    CellValue(final T value) {
        this.value = value;
    }

    public boolean isEmpty() {
        return value == null;
    }
}
