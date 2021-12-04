package net.summer.fastexcel.stream.utils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class TempFileUtil {
    private static final Logger LOGGER = LoggerFactory.getLogger(TempFileUtil.class);

    public static void quietDelete(final File file) {
        if (file != null) {
            LOGGER.debug("Deleting tmp file [" + file.getAbsolutePath() + "]");
            file.delete();
        }
    }

    public static File writeXlsxInputStreamToFile(final InputStream is) throws IOException {
        File f = Files.createTempFile("tmp-", ".xlsx").toFile();
        try (is; OutputStream fos = new FileOutputStream(f)) {
            is.transferTo(fos);
        }

        return f;
    }
}