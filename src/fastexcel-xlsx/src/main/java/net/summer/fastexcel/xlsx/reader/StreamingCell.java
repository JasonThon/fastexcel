package net.summer.fastexcel.xlsx.reader;

import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.Calendar;
import java.util.Date;
import java.util.Objects;
import java.util.Optional;

import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.ExcelNumberFormat;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import net.summer.fastexcel.stream.exceptions.NotSupportedException;
import net.summer.fastexcel.stream.utils.DateTimeUtils;

import static org.apache.poi.ss.usermodel.DateUtil.isADateFormat;

public class StreamingCell implements Cell {

    private static final Supplier NULL_SUPPLIER = () -> null;
    private static final String FALSE_AS_STRING = "0";
    private static final String TRUE_AS_STRING = "1";
    private boolean use1904Dates;
    private Sheet sheet;
    private int columnIndex;
    private int rowIndex;
    private Supplier contentsSupplier = NULL_SUPPLIER;
    private Object rawContents;
    private String formula;
    private String type;
    private CellStyle cellStyle;
    private Row row;
    private boolean formulaType;

    public StreamingCell(final Sheet sheet,
                         final int columnIndex,
                         final int rowIndex,
                         final boolean use1904Dates) {
        this.sheet = sheet;
        this.columnIndex = columnIndex;
        this.rowIndex = rowIndex;
        this.use1904Dates = use1904Dates;
    }

    public StreamingCell copy() {
        StreamingCell streamingCell = new StreamingCell(this.sheet, this.columnIndex, this.rowIndex, this.use1904Dates);
        streamingCell.setCellStyle(cellStyle);
        streamingCell.setRawContents(rawContents);
        streamingCell.setContentSupplier(contentsSupplier);
        streamingCell.setType(type);

        return streamingCell;
    }

    /**
     * Not supported
     */
    @Override
    public CellAddress getAddress() {
        throw new NotSupportedException();
    }

    /**
     * Not supported
     */
    @Override
    public CellRangeAddress getArrayFormulaRange() {
        throw new NotSupportedException();
    }

    /**
     * Get the value of the cell as a boolean. For strings we throw an exception. For
     * blank cells we return a false.
     *
     * @return the value of the cell as a date
     */
    @Override
    public boolean getBooleanCellValue() {
        CellType cellType = getCellType();
        switch (cellType) {
            case BLANK:
                return false;
            case BOOLEAN:
                return TRUE_AS_STRING.equals(rawContents);
            case FORMULA:
                throw new NotSupportedException();
            default:
                throw typeMismatch(CellType.BOOLEAN, cellType, false);
        }
    }

    /**
     * Only valid for formula cells
     *
     * @return one of ({@link CellType#NUMERIC}, {@link CellType#STRING},
     * {@link CellType#BOOLEAN}, {@link CellType#ERROR}) depending
     * on the cached value of the formula
     */
    @Override
    public CellType getCachedFormulaResultType() {
        throw new UnsupportedOperationException();
    }

    @Override
    public CellType getCachedFormulaResultTypeEnum() {
        return null;
    }

    /**
     * Not supported
     */
    @Override
    public Comment getCellComment() {
        throw new NotSupportedException();
    }

    /**
     * Not supported
     */
    @Override
    public void setCellComment(final Comment comment) {
        throw new NotSupportedException();
    }

    /**
     * Return a formula for the cell, for example, <code>SUM(C4:E4)</code>
     *
     * @return a formula for the cell
     * @throws IllegalStateException if the cell type returned by {@link #getCellType()} is not CELL_TYPE_FORMULA
     */
    @Override
    public String getCellFormula() {
        if (!formulaType) {
            throw new IllegalStateException("This cell does not have a formula");
        }

        return formula;
    }

    /**
     * Not supported
     */
    @Override
    public void setCellFormula(final String formula) throws FormulaParseException {
        throw new NotSupportedException();
    }

    /**
     * @return the style of the cell
     */
    @Override
    public CellStyle getCellStyle() {
        return this.cellStyle;
    }

    @Override
    public void setCellStyle(final CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    /**
     * Return the cell type.
     *
     * @return the cell type
     */
    @Override
    public CellType getCellType() {
        if (formulaType) {
            return CellType.FORMULA;
        } else if (contentsSupplier.getContent() == null || type == null) {
            return CellType.BLANK;
        } else if ("n".equals(type)) {
            return CellType.NUMERIC;
        } else if ("s".equals(type) || "inlineStr".equals(type) || "str".equals(type)) {
            return CellType.STRING;
        } else if ("str".equals(type)) {
            return CellType.FORMULA;
        } else if ("b".equals(type)) {
            return CellType.BOOLEAN;
        } else if ("e".equals(type)) {
            return CellType.ERROR;
        } else {
            throw new UnsupportedOperationException("Unsupported cell type '" + type + "'");
        }
    }

    /**
     * Not supported
     */
    @Override
    public void setCellType(final CellType cellType) {
        throw new NotSupportedException();
    }

    @Override
    public CellType getCellTypeEnum() {
        return getCellType();
    }

    /**
     * Returns column index of this cell
     *
     * @return zero-based column index of a column in a sheet.
     */
    @Override
    public int getColumnIndex() {
        return columnIndex;
    }

    public StreamingCell setColumnIndex(final int columnIndex) {
        this.columnIndex = columnIndex;
        return this;
    }

    /**
     * Get the value of the cell as a date. For strings we throw an exception. For
     * blank cells we return a null.
     *
     * @return the value of the cell as a date
     * @throws IllegalStateException if the cell type returned by {@link #getCellType()} is CELL_TYPE_STRING
     * @throws NumberFormatException if the cell value isn't a parsable <code>double</code>.
     */
    @Override
    public Date getDateCellValue() {
        if (getCellType() == CellType.STRING) {
            throw new IllegalStateException("Cell type cannot be CELL_TYPE_STRING");
        }

        return Optional.ofNullable(rawContents)
                .map(content -> DateUtil.getJavaDate(getNumericCellValue(), use1904Dates))
                .orElse(null);
    }

    /**
     * Not supported
     */
    @Override
    public byte getErrorCellValue() {
        throw new NotSupportedException();
    }

    /**
     * Not supported
     */
    @Override
    public Hyperlink getHyperlink() {
        throw new NotSupportedException();
    }

    /* Supported */

    /**
     * Not supported
     */
    @Override
    public void setHyperlink(final Hyperlink link) {
        throw new NotSupportedException();
    }

    @Override
    public LocalDateTime getLocalDateTimeCellValue() {
        return LocalDateTime.ofInstant(Instant.ofEpochMilli(getDateCellValue().getTime()), ZoneOffset.systemDefault());
    }

    /**
     * Get the value of the cell as a number. For strings we throw an exception. For
     * blank cells we return a 0.
     *
     * @return the value of the cell as a number
     * @throws NumberFormatException if the cell value isn't a parsable <code>double</code>.
     */
    @Override
    public double getNumericCellValue() {
        return Optional.ofNullable(rawContents)
                .map(contant -> Double.parseDouble(Objects.toString(contant)))
                .orElse(0.0);
    }

    /**
     * Get the value of the cell as a XSSFRichTextString
     * <p>
     * For numeric cells we throw an exception. For blank cells we return an empty string.
     * For formula cells we return the pre-calculated value if a string, otherwise an exception
     * </p>
     *
     * @return the value of the cell as a XSSFRichTextString
     */
    @Override
    public XSSFRichTextString getRichStringCellValue() {
        CellType cellType = getCellType();
        XSSFRichTextString rt;
        switch (cellType) {
            case BLANK:
                rt = new XSSFRichTextString("");
                break;
            case STRING:
                rt = new XSSFRichTextString(getStringCellValue());
                break;
            default:
                throw new NotSupportedException();
        }
        return rt;
    }

    /**
     * Returns the Row this cell belongs to. Note that keeping references to cell
     * rows around after the iterator window has passed <b>will</b> preserve them.
     *
     * @return the Row that owns this cell
     */
    @Override
    public Row getRow() {
        return row;
    }

    /**
     * Sets the Row this cell belongs to. Note that keeping references to cell
     * rows around after the iterator window has passed <b>will</b> preserve them.
     * <p>
     * The row is not automatically set.
     *
     * @param row The row
     */
    public void setRow(final Row row) {
        this.row = row;
    }

    /**
     * Returns row index of a row in the sheet that contains this cell
     *
     * @return zero-based row index of a row in the sheet that contains this cell
     */
    @Override
    public int getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(final int rowIndex) {
        this.rowIndex = rowIndex;
    }

    @Override
    public Sheet getSheet() {
        return sheet;
    }

    public void setSheet(final Sheet sheet) {
        this.sheet = sheet;
    }

    /**
     * Get the value of the cell as a string.
     * For blank cells we return an empty string.
     *
     * @return the value of the cell as a string
     */
    @Override
    public String getStringCellValue() {
        if (this.getCellType() == CellType.NUMERIC && isCellDateFormatted(this)) {
            return DateTimeUtils.toLocal(this.getDateCellValue().toInstant())
                    .format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
        }

        return Optional.ofNullable(rawContents)
                .map(Object::toString)
                .orElse("");
    }

    public String getType() {
        return type;
    }

    public void setType(final String type) {
        this.type = type;
    }

    public void init() {
        this.cellStyle = null;
        this.contentsSupplier = NULL_SUPPLIER;
        this.formulaType = false;
        this.row = null;
        this.rawContents = null;
        this.type = null;
        this.formula = null;
    }

    /**
     * Not supported
     */
    @Override
    public boolean isPartOfArrayFormulaGroup() {
        throw new NotSupportedException();
    }

    /**
     * Not supported
     */
    @Override
    public void removeCellComment() {
        throw new NotSupportedException();
    }

    /* Not supported */

    /**
     * Not supported
     */
    @Override
    public void removeFormula() throws IllegalStateException {
        throw new NotSupportedException();
    }

    /**
     * Not supported
     */
    @Override
    public void removeHyperlink() {
        throw new NotSupportedException();
    }

    /**
     * Not supported
     */
    @Override
    public void setAsActiveCell() {
        throw new NotSupportedException();
    }

    /**
     * Not supported
     */
    @Override
    public void setBlank() {
        throw new NotSupportedException();
    }

    /**
     * Not supported
     */
    @Override
    public void setCellErrorValue(final byte value) {
        throw new NotSupportedException();
    }

    /**
     * Not supported
     */
    @Override
    public void setCellValue(final double value) {
        throw new NotSupportedException();
    }

    /**
     * Not supported
     */
    @Override
    public void setCellValue(final Date value) {
        throw new NotSupportedException();
    }

    /**
     * Not supported
     */
    @Override
    public void setCellValue(final LocalDateTime value) {
        throw new NotSupportedException();
    }

    /**
     * Not supported
     */
    @Override
    public void setCellValue(final Calendar value) {
        throw new NotSupportedException();
    }

    /**
     * Not supported
     */
    @Override
    public void setCellValue(final RichTextString value) {
        throw new NotSupportedException();
    }

    /**
     * Not supported
     */
    @Override
    public void setCellValue(final String value) {
        throw new NotSupportedException();
    }

    /**
     * Not supported
     */
    @Override
    public void setCellValue(final boolean value) {
        throw new NotSupportedException();
    }

    public void setContentSupplier(final Supplier contentsSupplier) {
        this.contentsSupplier = contentsSupplier;
    }

    public void setFormula(final String formula) {
        this.formula = formula;
    }

    public void setFormulaType(final boolean formulaType) {
        this.formulaType = formulaType;
    }

    public void setRawContents(final Object rawContents) {
        this.rawContents = rawContents;
    }

    public StreamingCell setUse1904Dates(final boolean use1904Dates) {
        this.use1904Dates = use1904Dates;
        return this;
    }

    /**
     * Used to help format error messages
     */
    private static String getCellTypeName(final CellType cellType) {
        switch (cellType) {
            case BLANK:
                return "blank";
            case STRING:
                return "text";
            case BOOLEAN:
                return "boolean";
            case ERROR:
                return "error";
            case NUMERIC:
                return "numeric";
            case FORMULA:
                return "formula";
            default:
                return "#unknown cell type (" + cellType + ")#";
        }
    }

    private static RuntimeException typeMismatch(final CellType expectedType,
                                                 final CellType actualType,
                                                 final boolean isFormulaCell) {
        java.util.function.Supplier<String> getFormulaMsg = () -> {
            if (isFormulaCell) {
                return "formula ";
            }

            return "";
        };

        String builder = "Cannot get a "
                + getCellTypeName(expectedType)
                + " value from a "
                + getCellTypeName(actualType)
                + getFormulaMsg.get()
                + "cell";

        return new IllegalStateException(builder);
    }

    private boolean isCellDateFormatted(final StreamingCell cell) {
        if (DateUtil.isValidExcelDate(cell.getNumericCellValue())) {
            ExcelNumberFormat nf = ExcelNumberFormat.from(cell.cellStyle);
            if (nf == null) {
                return false;
            }
            return isADateFormat(nf);
        }
        return false;
    }
}