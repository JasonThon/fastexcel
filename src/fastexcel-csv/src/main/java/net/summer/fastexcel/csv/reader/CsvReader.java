package net.summer.fastexcel.csv.reader;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Iterator;

import net.summer.fastexcel.stream.MmapInputStream;

/**
 * @author songyi
 * @Date 2021/12/1
 **/
public class CsvReader implements Iterator<Record> {
    private final Iterator<String> iterator;
    private final Record record = new Record();
    private File file;

    public CsvReader(final InputStream stream) {
        if (stream instanceof MmapInputStream) {
            this.iterator = new BufferedReader(new InputStreamReader(stream))
                    .lines()
                    .iterator();
        } else {

            try {
                Path path = Files.createTempFile("csv-", ".tmp");
                this.file = path.toFile();
                try (OutputStream outputStream = new FileOutputStream(file)) {
                    stream.transferTo(outputStream);
                }

                this.iterator = new BufferedReader(new FileReader(file))
                        .lines()
                        .iterator();
            } catch (IOException e) {
                throw new RuntimeException(e);
            } finally {
                if (file != null) {
                    file.delete();
                }
            }
        }
    }

    public void close() {
        if (this.file != null) {
            this.file.delete();
        }
    }

    @Override
    public boolean hasNext() {
        return iterator.hasNext();
    }

    @Override
    public Record next() {
        return record.newValues(iterator.next());
    }
}
