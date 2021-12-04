package net.summer.fastexcel.xlsx.reader;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.util.IOUtils;
import org.xml.sax.Attributes;
import org.xml.sax.helpers.DefaultHandler;

import com.google.common.collect.Lists;

import net.summer.fastexcel.stream.utils.XmlUtils;

/**
 * @author songyi
 * @Date 2021/12/1
 **/
class LazySharedStringsTable implements AutoCloseable {
    private final SharedStringsTableHandler handler = new SharedStringsTableHandler();
    private InputStream inputStream;
    private List<String> strings;

    LazySharedStringsTable(final InputStream shareStringStream) {
        this.inputStream = shareStringStream;
    }

    @Override
    public void close() {
        IOUtils.closeQuietly(this.inputStream);

        this.strings.clear();
        this.handler.init();
    }

    public String getItemAt(final int idx) {
        if (strings == null || strings.size() == 0) {
            try {
                readFrom();
            } catch (IOException ignored) {
            }
        }

        return this.strings.get(idx);
    }

    private void readFrom() throws IOException {
        if (this.inputStream == null) {
            throw new RuntimeException("unable to parse shared strings table because of null input stream");
        }

        XmlUtils.parseXmlSource(this.inputStream, handler);

        this.strings = handler.getStrings();

        this.inputStream.close();
        this.inputStream = null;
    }

    static class SharedStringsTableHandler extends DefaultHandler {
        private static final String T_TAG = "t";
        private static final String SI_TAG = "si";
        private static final String RPH_TAG = "rPh";
        private final List<String> strings = Lists.newArrayList();
        private StringBuilder currentData;
        private StringBuilder currentElementData;
        private boolean ignoreTagt = false;
        private boolean isTagt = false;

        @Override
        public void characters(final char[] ch, final int start, final int length) {
            if (!isTagt || ignoreTagt) {
                return;
            }
            if (currentElementData == null) {
                currentElementData = new StringBuilder();
            }
            currentElementData.append(ch, start, length);
        }

        @Override
        public void endElement(final String uri, final String localName, final String qName) {
            if (T_TAG.equals(qName)) {
                if (currentElementData != null) {
                    if (currentData == null) {
                        currentData = new StringBuilder();
                    }
                    currentData.append(currentElementData);
                }
                isTagt = false;
            } else if (SI_TAG.equals(qName)) {
                if (currentData == null) {
                    strings.add(null);
                } else {
                    strings.add(currentData.toString());
                }
            } else if (RPH_TAG.equals(qName)) {
                ignoreTagt = false;
            }
        }

        public List<String> getStrings() {
            return strings;
        }

        public void init() {
            this.strings.clear();
            this.currentData = null;
            this.currentElementData = null;
        }

        @Override
        public void startElement(final String uri, final String localName,
                                 final String qName,
                                 final Attributes attributes) {
            if (T_TAG.equals(qName)) {
                currentElementData = null;
                isTagt = true;
            } else if (SI_TAG.equals(qName)) {
                currentData = null;
            } else if (RPH_TAG.equals(qName)) {
                ignoreTagt = true;
            }
        }
    }
}
