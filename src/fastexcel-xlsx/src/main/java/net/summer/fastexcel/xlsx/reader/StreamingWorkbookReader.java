package net.summer.fastexcel.xlsx.reader;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.channels.FileChannel;
import java.nio.file.StandardOpenOption;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;
import javax.xml.stream.XMLStreamException;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import net.summer.fastexcel.stream.MmapInputStream;
import net.summer.fastexcel.stream.exceptions.OpenException;
import net.summer.fastexcel.stream.exceptions.ReadException;
import net.summer.fastexcel.stream.utils.CollectionUtils;
import net.summer.fastexcel.stream.utils.TempFileUtil;
import net.summer.fastexcel.xlsx.reader.StreamingReader.Builder;

import static net.summer.fastexcel.stream.utils.TempFileUtil.writeXlsxInputStreamToFile;

public class StreamingWorkbookReader implements Iterable<Sheet>, AutoCloseable {
    private static final Logger LOGGER = LoggerFactory.getLogger(StreamingWorkbookReader.class);

    private final List<StreamingSheet> sheets;
    private final Builder builder;
    private boolean use1904Dates;
    private File tmp;
    private OPCPackage pkg;
    private FileChannel channel;

    public StreamingWorkbookReader(final Builder builder) {
        this.sheets = new ArrayList<>();
        this.builder = builder;
    }

    @Override
    public void close() {
        for (StreamingSheet sheet : sheets) {
            if (sheet.getReader() != null) {
                sheet.getReader().close();
            }
        }
        pkg.revert();

        IOUtils.closeQuietly(this.channel);
        TempFileUtil.quietDelete(tmp);
    }

    public int findSheetByName(final String name) {
        return 0;
    }

    public String findSheetName(final int sheetname) {
        StreamingSheet sheet = CollectionUtils.safeGet(sheets, sheetname);
        if (sheet != null) {

            return sheet.getSheetName();
        }
        return null;
    }

    public StreamingSheetReader first() {
        return sheets.get(0).getReader();
    }

    public void init(final InputStream is) {
        File f = null;
        try {
            f = writeXlsxInputStreamToFile(is);
            LOGGER.debug("Created temp file [" + f.getAbsolutePath() + "]");

            init(f);
            tmp = f;
        } catch (IOException e) {
            throw new ReadException("Unable to read input stream", e);
        } catch (RuntimeException e) {
            if (f != null) {
                f.delete();
            }
            throw e;
        }
    }

    public void init(final File f) {
        builder.check();
        try {
            if (builder.getPassword() != null) {
                // Based on: https://poi.apache.org/encryption.html
                this.channel = FileChannel.open(f.toPath(), StandardOpenOption.READ, StandardOpenOption.DELETE_ON_CLOSE);
                POIFSFileSystem poifs = new POIFSFileSystem(this.channel);
                EncryptionInfo info = new EncryptionInfo(poifs);
                Decryptor d = Decryptor.getInstance(info);
                d.verifyPassword(builder.getPassword());
                pkg = OPCPackage.open(d.getDataStream(poifs));
            } else {
                pkg = OPCPackage.open(f);
            }

            loadSheets(pkg, builder.getRowCacheSize());
        } catch (IOException e) {
            throw new OpenException("Failed to open file", e);
        } catch (OpenXML4JException | XMLStreamException e) {
            throw new ReadException("Unable to read workbook", e);
        } catch (GeneralSecurityException e) {
            throw new ReadException("Unable to read workbook - Decryption failed", e);
        }
    }

    @Override
    public Iterator<Sheet> iterator() {
        return new StreamingSheetIterator(sheets.iterator());
    }

    List<? extends Sheet> getSheets() {
        return sheets;
    }

    private void loadSheets(final OPCPackage pkg, final int rowCacheSize) throws IOException, OpenXML4JException,
            XMLStreamException {
        XSSFReader reader = new XSSFReader(pkg);
        List<PackagePart> parts = pkg.getPartsByContentType(XSSFRelation.SHARED_STRINGS.getContentType());
        InputStream sst = null;

        if (parts.size() > 0) {
            sst = parts.get(0).getInputStream();
        }

        SheetIterator iter = (SheetIterator) reader.getSheetsData();
        int index = 0;
        StyleTables styleTables = readTables(pkg);

        while (iter.hasNext()) {
            InputStream is = iter.next();
            String sheetName = iter.getSheetName();

            if (Objects.equals(builder.getSheetName(), sheetName)
                    || builder.getSheetIndex() == index) {

                sheets.add(
                        new StreamingSheet(
                                sheetName,
                                new StreamingSheetReader(
                                        sst,
                                        styleTables,
                                        new MmapInputStream(is, "sheet"),
                                        use1904Dates,
                                        rowCacheSize
                                )
                        )
                );
            } else {
                sheets.add(new StreamingSheet(sheetName, null));
            }

            index++;
            is.close();
        }
    }

    private StyleTables readTables(final OPCPackage pkg) throws IOException {
        List<PackagePart> parts = pkg.getPartsByContentType(XSSFRelation.STYLES.getContentType());
        if (parts.size() == 0) {
            return null;
        }

        return new StyleTables(parts.get(0));
    }

    static class StreamingSheetIterator implements Iterator<Sheet> {
        private final Iterator<StreamingSheet> iterator;

        StreamingSheetIterator(final Iterator<StreamingSheet> iterator) {
            this.iterator = iterator;
        }

        @Override
        public boolean hasNext() {
            return iterator.hasNext();
        }

        @Override
        public Sheet next() {
            return iterator.next();
        }

        @Override
        public void remove() {
            throw new RuntimeException("NotSupported");
        }
    }
}
