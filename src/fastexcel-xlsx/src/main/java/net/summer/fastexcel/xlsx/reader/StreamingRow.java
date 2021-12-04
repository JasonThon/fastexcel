package net.summer.fastexcel.xlsx.reader;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.google.common.collect.Lists;

import net.summer.fastexcel.stream.exceptions.NotSupportedException;
import net.summer.fastexcel.stream.utils.CollectionUtils;

public class StreamingRow implements Row {
    private Sheet sheet;
    private int rowIndex;
    private List<IndexedCell> cells;

    public StreamingRow(final Sheet sheet, final int rowIndex) {
        this.sheet = sheet;
        this.rowIndex = rowIndex;
    }

    public void addCell(final int columnIndex, final StreamingCell currentCell) {
        if (this.cells == null) {
            this.cells = new ArrayList<>();
        }

        this.cells.add(new IndexedCell(columnIndex, currentCell));
        this.cells.sort(Comparator.comparingInt(IndexedCell::getIndex));
    }

    /**
     * @return Cell iterator of the physically defined cells for this row.
     */
    @Override
    public Iterator<Cell> cellIterator() {
        if (cells == null) {
            return null;
        }

        return CollectionUtils.stream(cells)
                .map(IndexedCell::getCell)
                .iterator();
    }

    public void clearCell() {
        if (this.cells != null) {
            this.cells.clear();
        }

    }

    public Row copy() {
        StreamingRow row = new StreamingRow(this.sheet, rowIndex);
        row.cells = Lists.newArrayList(cells);

        return row;
    }

    /**
     * Not supported
     */
    @Override
    public Cell createCell(final int column) {
        throw new NotSupportedException();
    }

    /**
     * Not supported
     */
    @Override
    public Cell createCell(final int i, final CellType cellType) {
        throw new NotSupportedException();
    }

    /**
     * Get the cell representing a given column (logical cell) 0-based.  If you
     * ask for a cell that is not defined, you get a null.
     *
     * @param cellnum 0 based column number
     * @return Cell representing that column or null if undefined.
     */
    @Override
    public Cell getCell(final int cellnum) {
        return CollectionUtils.stream(cells)
                .filter(cell -> cellnum == cell.index)
                .findFirst()
                .map(IndexedCell::getCell)
                .orElse(null);
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Cell getCell(final int cellnum, final MissingCellPolicy policy) {
        StreamingCell cell = (StreamingCell) getCell(cellnum);
        if (policy == MissingCellPolicy.CREATE_NULL_AS_BLANK) {
            if (cell == null) {
                return new StreamingCell(sheet, cellnum, rowIndex, false);
            }
        } else if (policy == MissingCellPolicy.RETURN_BLANK_AS_NULL) {
            if (cell == null || cell.getCellType() == CellType.BLANK) {
                return null;
            }
        }

        return cell;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public short getFirstCellNum() {
        if (cells == null || cells.size() == 0) {
            return -1;
        }

        return cells.get(0).getIndex().shortValue();
    }

    /**
     * Not supported
     */
    @Override
    public short getHeight() {
        throw new NotSupportedException();
    }

    /**
     * Not supported
     */
    @Override
    public void setHeight(final short height) {
        throw new NotSupportedException();
    }

    /**
     * Not supported
     */
    @Override
    public float getHeightInPoints() {
        throw new NotSupportedException();
    }

    /**
     * Not supported
     */
    @Override
    public void setHeightInPoints(final float height) {
        throw new NotSupportedException();
    }

    /**
     * Gets the index of the last cell contained in this row <b>PLUS ONE</b>.
     *
     * @return short representing the last logical cell in the row <b>PLUS ONE</b>,
     * or -1 if the row does not contain any cells.
     */
    @Override
    public short getLastCellNum() {
        if (cells == null || cells.size() == 0) {
            return -1;
        }

        return cells.get(cells.size() - 1)
                .getIndex()
                .shortValue();
    }

    /**
     * Not supported
     */
    @Override
    public int getOutlineLevel() {
        throw new NotSupportedException();
    }

    /* Not supported */

    /**
     * Gets the number of defined cells (NOT number of cells in the actual row!).
     * That is to say if only columns 0,4,5 have values then there would be 3.
     *
     * @return int representing the number of defined cells in the row.
     */
    @Override
    public int getPhysicalNumberOfCells() {
        if (cells == null) {
            return 0;
        }

        return cells.size();
    }

    /**
     * Get row number this row represents
     *
     * @return the row number (0 based)
     */
    @Override
    public int getRowNum() {
        return rowIndex;
    }

    /**
     * Not supported
     */
    @Override
    public void setRowNum(final int rowNum) {
        this.rowIndex = rowNum;
    }

    /**
     * Not supported
     */
    @Override
    public CellStyle getRowStyle() {
        throw new NotSupportedException();
    }

    /**
     * Not supported
     */
    @Override
    public void setRowStyle(final CellStyle style) {
        throw new NotSupportedException();
    }

    @Override
    public Sheet getSheet() {
        return sheet;
    }

    public void setSheet(final Sheet sheet) {
        this.sheet = sheet;
    }

    /**
     * Get whether or not to display this row with 0 height
     *
     * @return - zHeight height is zero or not.
     */
    @Override
    public boolean getZeroHeight() {
        return false;
    }

    /**
     * Not supported
     */
    @Override
    public void setZeroHeight(final boolean zHeight) {
        throw new NotSupportedException();
    }

    /**
     * Not supported
     */
    @Override
    public boolean isFormatted() {
        throw new NotSupportedException();
    }

    /**
     * @return Cell iterator of the physically defined cells for this row.
     */
    @Override
    public Iterator<Cell> iterator() {
        if (cells == null) {
            return new ArrayList<Cell>().iterator();
        }

        return cells.stream()
                .sorted(Comparator.comparing(IndexedCell::getIndex))
                .map(IndexedCell::getCell)
                .iterator();
    }

    /**
     * Not supported
     */
    @Override
    public void removeCell(final Cell cell) {
        throw new NotSupportedException();
    }

    public void setStyleTables(final StyleTables styleTables) {
    }

    /**
     * Not supported
     */
    @Override
    public void shiftCellsLeft(final int firstShiftColumnIndex, final int lastShiftColumnIndex, final int step) {
        throw new NotSupportedException();
    }

    /**
     * Not supported
     */
    @Override
    public void shiftCellsRight(final int firstShiftColumnIndex, final int lastShiftColumnIndex, final int step) {
        throw new NotSupportedException();
    }

    static class IndexedCell {
        private Integer index;
        private Cell cell;

        IndexedCell(final Integer index, final Cell cell) {
            this.index = index;
            this.cell = cell;
        }

        public Cell getCell() {
            return cell;
        }

        public Integer getIndex() {
            return index;
        }
    }
}
