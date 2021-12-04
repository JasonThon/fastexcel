package net.summer.fastexcel.xlsx;

import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.google.common.base.Objects;
import com.google.common.collect.Lists;

import net.summer.fastexcel.stream.CellValue;
import net.summer.fastexcel.stream.RowStream;
import net.summer.fastexcel.stream.StringCell;
import net.summer.fastexcel.stream.exceptions.UnsupportedTypeException;

/**
 * @author songyi
 * @Date 2021/11/30
 **/
public class XlsxRowStream extends RowStream<String> {
    private Iterator<Cell> iterator;
    private CellValue<String> cellValue;
    private Row row;

    @Override
    public String[] asSlice() {
        List<String> results = Lists.newArrayList();

        forEachRemaining(value -> {
            if (!value.isEmpty()) {
                results.add(value.getValue());
            }
        });

        this.iterator = row.cellIterator();

        return results.toArray(String[]::new);
    }

    @Override
    public boolean hasNext() {
        return iterator.hasNext();
    }

    @Override
    public CellValue<String> next() {
        if (iterator.hasNext()) {
            String value = iterator.next().getStringCellValue();

            cellValue.setValue(value);
        }

        return cellValue;
    }

    @Override
    public <R> void row(final R rawData) {
        if (rawData instanceof Row) {
            this.row = (Row) rawData;
            this.iterator = this.row.cellIterator();
            this.cellValue = new StringCell(null);
        } else {
            throw new UnsupportedTypeException("Unsupported xlsx row data type");
        }
    }
}
