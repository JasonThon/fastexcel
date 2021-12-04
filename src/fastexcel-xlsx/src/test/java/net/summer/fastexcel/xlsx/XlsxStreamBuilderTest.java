package net.summer.fastexcel.xlsx;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.concurrent.atomic.AtomicInteger;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

/**
 * @author songyi
 * @Date 2021/12/4
 **/
class XlsxStreamBuilderTest {

    private InputStream stream = null;

    @BeforeEach
    void setUp() throws IOException {
        this.stream = new FileInputStream("src/test/resources/test-preview.xlsx");
    }

    @Test
    void testBuildReader() {
        XlsxStream stream = XlsxStreamBuilder.builder()
                .stream(this.stream)
                .build();
        AtomicInteger count = new AtomicInteger();

        stream.foreach(row -> count.getAndIncrement());
        stream.close();
    }
}