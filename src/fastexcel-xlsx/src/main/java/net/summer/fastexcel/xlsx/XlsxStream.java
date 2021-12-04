package net.summer.fastexcel.xlsx;

import java.util.Iterator;
import java.util.List;
import java.util.function.Consumer;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;

import com.google.common.collect.Lists;

import net.summer.fastexcel.stream.RowStream;
import net.summer.fastexcel.stream.Table2DStream;

/**
 * @author songyi
 * @Date 2021/12/4
 **/
public class XlsxStream extends Table2DStream<String> {
    private final XlsxRowStream rowStream;
    private final Workbook workbook;
    private final Sheet sheet;
    private Iterator<Row> rowIterator;

    public XlsxStream(final Workbook workbook,
                      final int sheetIdx,
                      final Integer limit) {
        super(limit);
        this.rowStream = new XlsxRowStream();
        this.workbook = workbook;
        this.sheet = workbook.getSheetAt(sheetIdx);
    }

    @Override
    public void close() {
        IOUtils.closeQuietly(rowStream);
        IOUtils.closeQuietly(workbook);
    }

    @Override
    public void foreach(final Consumer<RowStream<String>> consumer) {
        super.foreach(consumer);
        this.rowIterator = sheet.rowIterator();
    }

    @Override
    public boolean hasNext() {
        return rowIterator.hasNext();
    }

    @Override
    public XlsxRowStream next() {
        Row row = rowIterator.next();
        rowStream.row(row);

        return rowStream;
    }

    @Override
    public List<String> tableNames() {
        List<String> sheetNames = Lists.newArrayList();

        for (int idx = 0; idx < workbook.getNumberOfSheets(); idx++) {
            sheetNames.add(workbook.getSheetName(idx));
        }

        return sheetNames;
    }
}
