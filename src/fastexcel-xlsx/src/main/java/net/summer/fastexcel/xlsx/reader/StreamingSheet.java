package net.summer.fastexcel.xlsx.reader;

import java.util.Collection;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.AutoFilter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellRange;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PaneInformation;

public class StreamingSheet implements Sheet {

    private final String name;
    private final StreamingSheetReader reader;

    public StreamingSheet(final String name, final StreamingSheetReader reader) {
        this.name = name;
        this.reader = reader;
        if (reader != null) {
            reader.setSheet(this);
        }
    }

    /**
     * Not supported
     */
    @Override
    public int addMergedRegion(final CellRangeAddress region) {
        throw new UnsupportedOperationException();
    }

    /* Supported */

    /**
     * Not supported
     */
    @Override
    public int addMergedRegionUnsafe(final CellRangeAddress cellRangeAddress) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void addValidationData(final DataValidation dataValidation) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void autoSizeColumn(final int column) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void autoSizeColumn(final int column, final boolean useMergedCells) {
        throw new UnsupportedOperationException();
    }

    /* Unsupported */

    /**
     * Not supported
     */
    @Override
    public Drawing createDrawingPatriarch() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void createFreezePane(final int colSplit, final int rowSplit, final int leftmostColumn, final int topRow) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void createFreezePane(final int colSplit, final int rowSplit) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public Row createRow(final int rownum) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void createSplitPane(final int xSplitPos, final int ySplitPos, final int leftmostColumn, final int topRow,
                                final int activePane) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public CellAddress getActiveCell() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setActiveCell(final CellAddress cellAddress) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public boolean getAutobreaks() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setAutobreaks(final boolean value) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public Comment getCellComment(final CellAddress cellAddress) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public Map<CellAddress, ? extends Comment> getCellComments() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public int[] getColumnBreaks() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public int getColumnOutlineLevel(final int columnIndex) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public CellStyle getColumnStyle(final int column) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public int getColumnWidth(final int columnIndex) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public float getColumnWidthInPixels(final int columnIndex) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public DataValidationHelper getDataValidationHelper() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public List<? extends DataValidation> getDataValidations() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public int getDefaultColumnWidth() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setDefaultColumnWidth(final int width) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public short getDefaultRowHeight() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setDefaultRowHeight(final short height) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public float getDefaultRowHeightInPoints() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setDefaultRowHeightInPoints(final float height) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public boolean getDisplayGuts() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setDisplayGuts(final boolean value) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public Drawing getDrawingPatriarch() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public int getFirstRowNum() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public boolean getFitToPage() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setFitToPage(final boolean value) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public Footer getFooter() {
        throw new UnsupportedOperationException();
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
    public Header getHeader() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public boolean getHorizontallyCenter() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setHorizontallyCenter(final boolean value) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public Hyperlink getHyperlink(final int i, final int i1) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public Hyperlink getHyperlink(final CellAddress cellAddress) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public List<? extends Hyperlink> getHyperlinkList() {
        throw new UnsupportedOperationException();
    }

    /**
     * Gets the last row on the sheet
     *
     * @return last row contained n this sheet (0-based)
     */
    @Override
    public int getLastRowNum() {
        return reader.getLastRowNum();
    }

    /**
     * Not supported
     */
    @Override
    public short getLeftCol() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public double getMargin(final short margin) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public CellRangeAddress getMergedRegion(final int index) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public List<CellRangeAddress> getMergedRegions() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public int getNumMergedRegions() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public PaneInformation getPaneInformation() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public int getPhysicalNumberOfRows() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public PrintSetup getPrintSetup() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public boolean getProtect() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public CellRangeAddress getRepeatingColumns() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setRepeatingColumns(final CellRangeAddress columnRangeRef) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public CellRangeAddress getRepeatingRows() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setRepeatingRows(final CellRangeAddress rowRangeRef) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public Row getRow(final int rownum) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public int[] getRowBreaks() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public boolean getRowSumsBelow() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setRowSumsBelow(final boolean value) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public boolean getRowSumsRight() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setRowSumsRight(final boolean value) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public boolean getScenarioProtect() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public SheetConditionalFormatting getSheetConditionalFormatting() {
        throw new UnsupportedOperationException();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getSheetName() {
        return name;
    }

    /**
     * Not supported
     */
    @Override
    public short getTopRow() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public boolean getVerticallyCenter() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setVerticallyCenter(final boolean value) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public Workbook getWorkbook() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void groupColumn(final int fromColumn, final int toColumn) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void groupRow(final int fromRow, final int toRow) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public boolean isColumnBroken(final int column) {
        throw new UnsupportedOperationException();
    }

    /**
     * Get the hidden state for a given column
     *
     * @param columnIndex - the column to set (0-based)
     * @return hidden - <code>false</code> if the column is visible
     */
    @Override
    public boolean isColumnHidden(final int columnIndex) {
        return false;
    }

    /**
     * Not supported
     */
    @Override
    public boolean isDisplayFormulas() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setDisplayFormulas(final boolean show) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public boolean isDisplayGridlines() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setDisplayGridlines(final boolean show) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public boolean isDisplayRowColHeadings() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setDisplayRowColHeadings(final boolean show) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public boolean isDisplayZeros() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setDisplayZeros(final boolean value) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public boolean isPrintGridlines() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setPrintGridlines(final boolean show) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public boolean isPrintRowAndColumnHeadings() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setPrintRowAndColumnHeadings(final boolean b) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public boolean isRightToLeft() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setRightToLeft(final boolean value) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public boolean isRowBroken(final int row) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public boolean isSelected() {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setSelected(final boolean value) {
        throw new UnsupportedOperationException();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Iterator<Row> iterator() {
        return reader.iterator();
    }

    /**
     * Not supported
     */
    @Override
    public void protectSheet(final String password) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public CellRange<? extends Cell> removeArrayFormula(final Cell cell) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void removeColumnBreak(final int column) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void removeMergedRegion(final int index) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void removeMergedRegions(final Collection<Integer> collection) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void removeRow(final Row row) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void removeRowBreak(final int row) {
        throw new UnsupportedOperationException();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Iterator<Row> rowIterator() {
        return reader.iterator();
    }

    /**
     * Not supported
     */
    @Override
    public CellRange<? extends Cell> setArrayFormula(final String formula, final CellRangeAddress range) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public AutoFilter setAutoFilter(final CellRangeAddress range) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setColumnBreak(final int column) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setColumnGroupCollapsed(final int columnNumber, final boolean collapsed) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setColumnHidden(final int columnIndex, final boolean hidden) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setColumnWidth(final int columnIndex, final int width) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setDefaultColumnStyle(final int column, final CellStyle style) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setMargin(final short margin, final double size) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setRowBreak(final int row) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setRowGroupCollapsed(final int row, final boolean collapse) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void setZoom(final int i) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void shiftColumns(final int startColumn, final int endColumn, final int n) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void shiftRows(final int startRow, final int endRow, final int n) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void shiftRows(final int startRow, final int endRow, final int n, final boolean copyRowHeight,
                          final boolean resetOriginalRowHeight) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void showInPane(final int toprow, final int leftcol) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void ungroupColumn(final int fromColumn, final int toColumn) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void ungroupRow(final int fromRow, final int toRow) {
        throw new UnsupportedOperationException();
    }

    /**
     * Not supported
     */
    @Override
    public void validateMergedRegions() {
        throw new UnsupportedOperationException();
    }

    StreamingSheetReader getReader() {
        return reader;
    }
}
