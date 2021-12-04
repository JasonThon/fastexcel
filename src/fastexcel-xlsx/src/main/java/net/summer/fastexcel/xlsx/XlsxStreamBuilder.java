package net.summer.fastexcel.xlsx;

import java.io.BufferedInputStream;
import java.io.InputStream;
import java.util.Objects;

import org.apache.poi.ss.usermodel.Workbook;

import net.summer.fastexcel.stream.Table2DStreamBuilder;
import net.summer.fastexcel.xlsx.reader.StreamingReader;

/**
 * @author songyi
 * @Date 2021/12/4
 **/
public class XlsxStreamBuilder implements Table2DStreamBuilder<String> {
    private Integer limit;
    private String sheetName;
    private int sheetIdx;
    private String password;
    private InputStream stream;

    public static XlsxStreamBuilder builder() {
        return new XlsxStreamBuilder();
    }

    @Override
    public XlsxStream build() {
        this.check();

        Workbook workbook = StreamingReader.builder()
                .sheetName(sheetName)
                .sheetIndex(sheetIdx)
                .password(password)
                .open(stream);

        return new XlsxStream(workbook, sheetIdx, limit);
    }

    public XlsxStreamBuilder limit(final int limit) {
        this.limit = limit;
        return this;
    }

    public XlsxStreamBuilder password(final String pwd) {
        this.password = pwd;
        return this;
    }

    public XlsxStreamBuilder sheetIdx(final int idx) {
        this.sheetIdx = idx;
        return this;
    }

    public XlsxStreamBuilder sheetName(final String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    public XlsxStreamBuilder stream(final InputStream stream) {
        if (stream instanceof BufferedInputStream) {
            this.stream = stream;
        } else {
            this.stream = new BufferedInputStream(stream);
        }

        return this;
    }

    private void check() {
        Objects.requireNonNull(stream, "xlsx file Stream should not be empty");
    }
}
