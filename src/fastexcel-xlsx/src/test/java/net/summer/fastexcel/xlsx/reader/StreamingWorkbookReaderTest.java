package net.summer.fastexcel.xlsx.reader;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * @author songyi
 * @Date 2021/12/3
 **/
class StreamingWorkbookReaderTest {

    private InputStream stream = null;

    @BeforeEach
    void setUp() throws IOException {
        this.stream = new FileInputStream("src/test/resources/test-preview.xlsx");
    }

    @Test
    void testRead() throws IOException {
        Workbook workbook = new StreamingReader.Builder()
                .rowCacheSize(20)
                .sheetIndex(0)
                .open(stream);

        Sheet sheet = workbook.getSheetAt(0);

        int count = 0;
        int columnCount = 0;

        for (Row row : sheet) {
            for (Cell cell : row) {
                if (count == 0) {
                    columnCount++;
                }
                // do nothing
            }
            count++;
        }

        workbook.close();

        assertEquals(32, count);
        assertEquals(3, columnCount);
    }
}