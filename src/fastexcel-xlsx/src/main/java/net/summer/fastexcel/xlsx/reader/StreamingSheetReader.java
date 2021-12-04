package net.summer.fastexcel.xlsx.reader;

import java.io.InputStream;
import java.util.Iterator;
import java.util.List;
import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.Characters;
import javax.xml.stream.events.EndElement;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.StaxHelper;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import com.google.common.collect.Lists;

import net.summer.fastexcel.stream.exceptions.CloseException;

public class StreamingSheetReader implements Iterable<Row> {

    private final LazySharedStringsTable sst;
    private final StyleTables styleTables;
    private final XMLEventReader parser;
    private final InputStream stream;
    private final boolean use1904Dates;
    private final int rowCacheSize;
    private List<Row> rowCache;
    private int lastRowNum;
    private int currentRowNum;
    private int firstColNum = 0;
    private int currentColNum;
    private Iterator<Row> rowCacheIterator;
    private String lastContents;
    private Sheet sheet;
    private StreamingRow currentRow;
    private StreamingCell currentCell;
    private boolean isEof;

    public StreamingSheetReader(final InputStream sst,
                                final StyleTables styleTables,
                                final InputStream stream,
                                final boolean use1904Dates,
                                final int rowCacheSize) throws XMLStreamException {
        this.sst = new LazySharedStringsTable(sst);
        this.styleTables = styleTables;
        this.parser = StaxHelper
                .newXMLInputFactory()
                .createXMLEventReader(stream);
        this.use1904Dates = use1904Dates;
        this.rowCacheSize = rowCacheSize;
        this.stream = stream;
    }

    public void close() {
        try {
            this.parser.close();
            this.sst.close();
            IOUtils.closeQuietly(this.stream);
        } catch (XMLStreamException e) {
            throw new CloseException(e);
        }
    }

    /**
     * Returns a new streaming iterator to loop through rows. This iterator is not
     * guaranteed to have all rows in memory, and any particular iteration may
     * trigger a load from disk to read in new data.
     *
     * @return the streaming iterator
     */
    @Override
    public Iterator<Row> iterator() {
        return new StreamingRowIterator();
    }

    /**
     * Tries to format the contents of the last contents appropriately based on
     * the type of cell and the discovered numeric format.
     *
     * @return
     */
    Supplier formattedContents() {
        return getFormatterForType(currentCell.getType());
    }

    /**
     * Gets the last row on the sheet
     *
     * @return
     */
    int getLastRowNum() {
        if (rowCacheIterator == null) {
            getRow();
        }
        return lastRowNum;
    }

    void setSheet(final StreamingSheet sheet) {
        this.sheet = sheet;
    }

    /**
     * Returns the contents of the cell, with no formatting applied
     *
     * @return
     */
    String unformattedContents() {
        switch (currentCell.getType()) {
            case "s":           //string stored in shared table
                if (!lastContents.isEmpty()) {
                    int idx = Integer.parseInt(lastContents);
                    return sst.getItemAt(idx);
                }
                return lastContents;
            case "inlineStr":   //inline string (not in sst)
                return new XSSFRichTextString(lastContents).toString();
            default:
                return lastContents;
        }
    }

    /**
     * Tries to format the contents of the last contents appropriately based on
     * the provided type and the discovered numeric format.
     *
     * @return
     */
    private Supplier getFormatterForType(final String type) {
        return new StringSupplier(lastContents);
    }

    /**
     * Read through a number of rows equal to the rowCacheSize field or until there is no more data to read
     *
     * @return true if data was read
     */
    private boolean getRow() {
        if (this.isEof) {
            return rowCacheIterator.hasNext();
        }

        if (this.rowCache == null) {
            this.rowCache = Lists.newArrayList();
        }

        this.rowCache.clear();

        try {
            while (this.rowCache.size() < rowCacheSize && parser.hasNext()) {
                handleEvent(parser.nextEvent());
            }
            rowCacheIterator = rowCache.iterator();
            return rowCacheIterator.hasNext();
        } catch (XMLStreamException e) {
            this.isEof = true;

            rowCacheIterator = rowCache.iterator();
            return rowCacheIterator.hasNext();
        }
    }

    /**
     * Handles a SAX event.
     *
     * @param event
     */
    private void handleEvent(final XMLEvent event) {
        if (event.getEventType() == XMLStreamConstants.CHARACTERS) {
            Characters c = event.asCharacters();
            lastContents += c.getData();
        } else if (event.getEventType() == XMLStreamConstants.START_ELEMENT
                && isSpreadsheetTag(event.asStartElement().getName())) {
            StartElement startElement = event.asStartElement();
            String tagLocalName = startElement.getName().getLocalPart();

            if ("row".equals(tagLocalName)) {
                Attribute rowNumAttr = startElement.getAttributeByName(new QName("r"));
                int rowIndex = currentRowNum;
                if (rowNumAttr != null) {
                    rowIndex = Integer.parseInt(rowNumAttr.getValue()) - 1;
                    currentRowNum = rowIndex;
                }
                if (currentRow == null) {
                    currentRow = new StreamingRow(sheet, rowIndex);
                } else {
                    currentRow.setRowNum(rowIndex);
                    currentRow.setSheet(sheet);
                    currentRow.clearCell();
                    currentRow.setStyleTables(styleTables);
                }
                currentColNum = firstColNum;
            } else if ("c".equals(tagLocalName)) {
                Attribute ref = startElement.getAttributeByName(new QName("r"));

                if (ref != null) {
                    String[] coord = splitCellRef(ref.getValue());
                    currentColNum = CellReference.convertColStringToIndex(coord[0]);
                    initCurrentCell(Integer.parseInt(coord[1]) - 1);
                } else {
                    initCurrentCell(currentRowNum);
                }
                Attribute style = startElement.getAttributeByName(new QName("s"));

                if (style != null) {
                    String indexStr = style.getValue();
                    int index = Integer.parseInt(indexStr);
                    currentCell.setCellStyle(styleTables.getIndexAt(index));
                }

                Attribute type = startElement.getAttributeByName(new QName("t"));
                if (type != null) {
                    currentCell.setType(type.getValue());
                } else {
                    currentCell.setType("n");
                }

            } else if ("dimension".equals(tagLocalName)) {
                Attribute refAttr = startElement.getAttributeByName(new QName("ref"));
                if (refAttr != null) {
                    String ref = refAttr.getValue();
                    // ref is formatted as A1 or A1:F25. Take the last numbers of this string and use it as lastRowNum
                    for (int i = ref.length() - 1; i >= 0; i--) {
                        if (!Character.isDigit(ref.charAt(i))) {
                            try {
                                lastRowNum = Integer.parseInt(ref.substring(i + 1)) - 1;
                            } catch (NumberFormatException ignore) {
                            }
                            break;
                        }
                    }
                    for (int i = 0; i < ref.length(); i++) {
                        if (!Character.isAlphabetic(ref.charAt(i))) {
                            firstColNum = CellReference.convertColStringToIndex(ref.substring(0, i));
                            break;
                        }
                    }
                }
            } else if ("f".equals(tagLocalName)) {
                if (currentCell != null) {
                    currentCell.setFormulaType(true);
                }
            }

            // Clear contents cache
            lastContents = "";
        } else if (event.getEventType() == XMLStreamConstants.END_ELEMENT
                && isSpreadsheetTag(event.asEndElement().getName())) {
            EndElement endElement = event.asEndElement();
            String tagLocalName = endElement.getName().getLocalPart();

            if ("v".equals(tagLocalName) || "t".equals(tagLocalName)) {
                currentCell.setRawContents(unformattedContents());
                currentCell.setContentSupplier(formattedContents());
            } else if ("row".equals(tagLocalName) && currentRow != null) {
                rowCache.add(currentRow.copy());
                currentRowNum++;
            } else if ("c".equals(tagLocalName)) {
                currentRow.addCell(currentCell.getColumnIndex(), currentCell.copy());
                currentCell = null;
                currentColNum++;
            } else if ("f".equals(tagLocalName)) {
                if (currentCell != null) {
                    currentCell.setFormula(lastContents);
                }
            }

        }
    }

    private void initCurrentCell(final int rowNum) {
        if (currentCell == null) {
            currentCell = new StreamingCell(sheet, currentColNum, rowNum, use1904Dates);
        } else {
            currentCell.init();
            currentCell.setSheet(sheet);
            currentCell.setColumnIndex(currentColNum);
            currentCell.setRowIndex(rowNum);
            currentCell.setUse1904Dates(use1904Dates);
        }
    }

    /**
     * Returns true if a tag is part of the main namespace for SpreadsheetML:
     * <ul>
     * <li>http://schemas.openxmlformats.org/spreadsheetml/2006/main
     * <li>http://purl.oclc.org/ooxml/spreadsheetml/main
     * </ul>
     * As opposed to http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing, etc.
     *
     * @param name
     * @return
     */
    private boolean isSpreadsheetTag(final QName name) {
        return (name.getNamespaceURI() != null
                && name.getNamespaceURI().endsWith("/main"));
    }

    private String[] splitCellRef(final String ref) {
        int splitPos = -1;

        // start at pos 1, since the first char is expected to always be a letter
        for (int i = 1; i < ref.length(); i++) {
            char c = ref.charAt(i);

            if (c >= '0' && c <= '9') {
                splitPos = i;
                break;
            }
        }

        return new String[]{
                ref.substring(0, splitPos),
                ref.substring(splitPos)
        };
    }

    class StreamingRowIterator implements Iterator<Row> {

        @Override
        public boolean hasNext() {
            if (rowCacheIterator == null || !rowCacheIterator.hasNext()) {
                return getRow();
            }

            return true;
        }

        @Override
        public Row next() {
            return rowCacheIterator.next();
        }

        @Override
        public void remove() {
            throw new RuntimeException("NotSupported");
        }
    }
}
