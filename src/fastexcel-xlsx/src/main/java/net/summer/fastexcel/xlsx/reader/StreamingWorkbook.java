package net.summer.fastexcel.xlsx.reader;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.udf.UDFFinder;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.ss.usermodel.Workbook;

import net.summer.fastexcel.stream.exceptions.MissingSheetException;
import net.summer.fastexcel.stream.utils.CollectionUtils;

public class StreamingWorkbook implements Workbook, AutoCloseable {
    private final StreamingWorkbookReader reader;

    public StreamingWorkbook(final StreamingWorkbookReader reader) {
        this.reader = reader;
    }

    /**
     * Not supported
     */
    @Override
    public int addOlePackage(final byte[] bytes, final String s, final String s1, final String s2) throws IOException {
        throw new UnsupportedOperationException();
    }

    /* Supported */

    /**
     * Not supported
     */
    @Override
    public int addPicture(final byte[] pictureData, final int format) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void addToolPack(final UDFFinder toopack) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public Sheet cloneSheet(final int sheetNum) {
        throw new UnsupportedOperationException();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void close() throws IOException {
        reader.close();
    }

    /**
     * Not supported
     */
    @Override
    public CellStyle createCellStyle() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public DataFormat createDataFormat() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public Font createFont() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public Name createName() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public Sheet createSheet() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public Sheet createSheet(final String sheetname) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public Font findFont(final boolean b, final short i, final short i1, final String s, final boolean b1, final boolean b2,
                         final short i2, final byte b3) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public int getActiveSheetIndex() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public List<? extends Name> getAllNames() {
        throw new UnsupportedOperationException();
    }

    /* Not supported */

    /**
     * Not supported
     */
    @Override
    public List<? extends PictureData> getAllPictures() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public CellStyle getCellStyleAt(final int i) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public CreationHelper getCreationHelper() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public int getFirstVisibleTab() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setFirstVisibleTab(final int sheetIndex) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public Font getFontAt(final int i) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Font getFontAt(final short i) {
        return null;
    }

    /**
     * Not supported
     */
    @Override
    public boolean getForceFormulaRecalculation() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setForceFormulaRecalculation(final boolean value) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public MissingCellPolicy getMissingCellPolicy() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setMissingCellPolicy(final MissingCellPolicy missingCellPolicy) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public Name getName(final String name) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Name getNameAt(final int i) {
        return null;
    }

    @Override
    public int getNameIndex(final String s) {
        return 0;
    }

    /**
     * Not supported
     */
    @Override
    public List<? extends Name> getNames(final String s) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public int getNumCellStyles() {
        throw new UnsupportedOperationException();
    }

    @Override
    public short getNumberOfFonts() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public int getNumberOfFontsAsInt() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public int getNumberOfNames() {
        throw new UnsupportedOperationException();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getNumberOfSheets() {
        return CollectionUtils.size(reader.getSheets());
    }

    /**
     * Not supported
     */
    @Override
    public String getPrintArea(final int sheetIndex) {
        throw new UnsupportedOperationException();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Sheet getSheet(final String name) {
        int index = getSheetIndex(name);
        if (index == -1) {
            throw new MissingSheetException("Sheet '" + name + "' does not exist");
        }
        return reader.getSheets().get(index);
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Sheet getSheetAt(final int index) {
        return reader.getSheets().get(index);
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getSheetIndex(final String name) {
        return findSheetByName(name);
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getSheetIndex(final Sheet sheet) {
        if (sheet instanceof StreamingSheet) {
            return findSheetByName(sheet.getSheetName());
        } else {
            throw new UnsupportedOperationException("Cannot use non-StreamingSheet sheets");
        }
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getSheetName(final int sheet) {
        return reader.findSheetName(sheet);
    }

    /**
     * Not supported
     */
    @Override
    public SheetVisibility getSheetVisibility(final int i) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public SpreadsheetVersion getSpreadsheetVersion() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public boolean isHidden() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setHidden(final boolean hiddenFlag) {
        throw new UnsupportedOperationException();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean isSheetHidden(final int sheetIx) {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean isSheetVeryHidden(final int sheetIx) {
        return false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Iterator<Sheet> iterator() {
        return reader.iterator();
    }

    /**
     * Not supported
     */
    @Override
    public int linkExternalWorkbook(final String name, final Workbook workbook) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void removeName(final int i) {

    }

    @Override
    public void removeName(final String s) {

    }

    /**
     * Not supported
     */
    @Override
    public void removeName(final Name name) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void removePrintArea(final int sheetIndex) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void removeSheetAt(final int index) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setActiveSheet(final int sheetIndex) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setPrintArea(final int sheetIndex, final String reference) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setPrintArea(final int sheetIndex, final int startColumn, final int endColumn, final int startRow,
                             final int endRow) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setSelectedTab(final int index) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setSheetHidden(final int sheetIx, final boolean hidden) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setSheetName(final int sheet, final String name) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setSheetOrder(final String sheetname, final int pos) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setSheetVisibility(final int i, final SheetVisibility sheetVisibility) {
        throw new UnsupportedOperationException();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Iterator<Sheet> sheetIterator() {
        return iterator();
    }

    /**
     * Not supported
     */
    @Override
    public void write(final OutputStream stream) {
        throw new UnsupportedOperationException();
    }

    int findSheetByName(final String name) {
        return reader.findSheetByName(name);

    }
}
