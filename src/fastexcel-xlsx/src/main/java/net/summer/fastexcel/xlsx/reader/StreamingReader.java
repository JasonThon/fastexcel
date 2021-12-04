package net.summer.fastexcel.xlsx.reader;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.Objects;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Streaming Excel workbook implementation. Most advanced features of POI are not supported.
 * Use this only if your application can handle iterating through an entire workbook, row by
 * row.
 */
public class StreamingReader implements Iterable<Row>, AutoCloseable {
    private final StreamingWorkbookReader workbook;

    public StreamingReader(final StreamingWorkbookReader workbook) {
        this.workbook = workbook;
    }

    public static Builder builder() {
        return new Builder();
    }

    @Override
    public void close() throws IOException {
        workbook.close();
    }

    @Override
    public Iterator<Row> iterator() {
        return workbook.first().iterator();
    }

    public static class Builder {
        private static final int DEFAULT_ROW_CACHE_SIZE = 10;

        private int rowCacheSize = DEFAULT_ROW_CACHE_SIZE;
        private int sheetIndex = 0;
        private String sheetName;
        private String password;

        public void check() {
            if (Objects.isNull(this.sheetName) && this.sheetIndex < 0) {
                throw new RuntimeException("Builder is not legal");
            }
        }

        /**
         * @return The password to use to unlock this workbook
         */
        public String getPassword() {
            return password;
        }

        public int getRowCacheSize() {
            return rowCacheSize;
        }

        public int getSheetIndex() {
            return sheetIndex;
        }

        public String getSheetName() {
            return sheetName;
        }

        public Workbook open(final InputStream is) {
            StreamingWorkbookReader workbook = new StreamingWorkbookReader(this);
            workbook.init(is);
            return new StreamingWorkbook(workbook);
        }

        public Workbook open(final File file) {
            StreamingWorkbookReader workbook = new StreamingWorkbookReader(this);
            workbook.init(file);
            return new StreamingWorkbook(workbook);
        }

        /**
         * For password protected files specify password to open file.
         * If the password is incorrect a {@code ReadException} is thrown on
         * {@code read}.
         * <p>NULL indicates that no password should be used, this is the
         * default value.</p>
         *
         * @param password to use when opening file
         * @return reference to current {@code Builder}
         */
        public Builder password(final String password) {
            this.password = password;
            return this;
        }

        public Builder rowCacheSize(final int rowCacheSize) {
            this.rowCacheSize = rowCacheSize;
            return this;
        }

        public Builder sheetIndex(final int sheetIndex) {
            this.sheetIndex = sheetIndex;
            return this;
        }

        public Builder sheetName(final String sheetName) {
            this.sheetName = sheetName;
            return this;
        }

    }

}
