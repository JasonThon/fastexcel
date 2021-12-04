package net.summer.fastexcel.stream;

import lombok.Getter;

/**
 * @author songyi
 * @Date 2021/12/3
 **/
public enum ExcelType {
    XLSX("xlsx"),
    CSV("csv"),
    XLS("xls");

    @Getter
    private final String value;

    ExcelType(final String value) {
        this.value = value;
    }
}
