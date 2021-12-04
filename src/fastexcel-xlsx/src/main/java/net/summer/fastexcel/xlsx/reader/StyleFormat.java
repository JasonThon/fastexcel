package net.summer.fastexcel.xlsx.reader;

import lombok.Getter;

/**
 * @author songyi
 * @Date 2021/12/2
 **/
@Getter
class StyleFormat extends PartialCellStyle {
    private final String dataFormatString;
    private short dataFormat;

    StyleFormat(final long dataFormat, final String dataFormatString) {
        this.dataFormat = (short) dataFormat;
        this.dataFormatString = dataFormatString;
    }

    @Override
    public void setDataFormat(final short fmt) {
        this.dataFormat = fmt;
    }
}
