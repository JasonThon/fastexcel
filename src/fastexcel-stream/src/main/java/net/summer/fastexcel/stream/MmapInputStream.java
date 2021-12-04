package net.summer.fastexcel.stream;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.MappedByteBuffer;
import java.nio.channels.FileChannel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;

/**
 * @author songyi
 * @Date 2021/11/30
 * <p>
 * Mmap InputStream. It will call mmap() to reduce memory copy and usage
 **/
public class MmapInputStream extends InputStream {
    private final FileChannel channel;
    private long offset;

    public MmapInputStream(final InputStream stream, final String prefix) throws IOException {
        Path path = Files.createTempFile(prefix, ".tmp");
        File tmp = path.toFile();

        try (OutputStream outputStream = new FileOutputStream(tmp)) {
            stream.transferTo(outputStream);
        }

        this.channel = openChannel(path);
    }

    @Override
    public void close() throws IOException {
        this.channel.close();
    }

    @Override
    public int read() throws IOException {
        MappedByteBuffer buffer = this.channel.map(FileChannel.MapMode.READ_ONLY, offset, 1L);

        if (buffer.hasRemaining()) {
            offset++;

            return buffer.get();
        }

        return -1;
    }

    @Override
    public int read(final byte[] b, final int off, final int len) throws IOException {
        MappedByteBuffer buffer = this.channel.map(FileChannel.MapMode.READ_ONLY, offset, len);

        if (!buffer.hasRemaining()) {
            return -1;
        }

        int position = buffer.get(b, off, len).position();
        offset += position;
        return position;
    }

    private FileChannel openChannel(final Path path) throws IOException {
        return FileChannel.open(
                path,
                StandardOpenOption.READ,
                StandardOpenOption.WRITE,
                StandardOpenOption.DELETE_ON_CLOSE
        );
    }
}
