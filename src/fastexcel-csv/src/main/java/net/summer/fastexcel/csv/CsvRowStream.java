package net.summer.fastexcel.csv;

import net.summer.fastexcel.csv.reader.Record;
import net.summer.fastexcel.stream.CellValue;
import net.summer.fastexcel.stream.RowStream;
import net.summer.fastexcel.stream.StringCell;

/**
 * @author songyi
 * @Date 2021/12/1
 **/
public class CsvRowStream extends RowStream<String> {
    private Record record;
    private int currentIndex;

    @Override
    public String[] asSlice() {
        return record.getValues()
                .toArray(String[]::new);
    }

    @Override
    public boolean hasNext() {
        return record.isIdxOutOfBound(currentIndex);
    }

    @Override
    public CellValue<String> next() {
        currentIndex++;
        return new StringCell(record.getItemAtIdx(currentIndex - 1));
    }

    @Override
    public <R> void row(final R rawData) {
        if (rawData instanceof Record) {
            this.record = (Record) rawData;
        } else {
            throw new UnsupportedOperationException("Not supported csv row data");
        }
    }
}
