package net.summer.fastexcel.xlsx.reader;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * @author songyi
 * @Date 2021/12/2
 **/
abstract class PartialCellStyle implements CellStyle {

    @Override
    public void cloneStyleFrom(final CellStyle source) {

    }

    @Override
    public HorizontalAlignment getAlignment() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setAlignment(final HorizontalAlignment align) {
        throw new UnsupportedOperationException();
    }

    @Override
    public HorizontalAlignment getAlignmentEnum() {
        throw new UnsupportedOperationException();
    }

    @Override
    public BorderStyle getBorderBottom() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setBorderBottom(final BorderStyle border) {
        throw new UnsupportedOperationException();
    }

    @Override
    public BorderStyle getBorderBottomEnum() {
        throw new UnsupportedOperationException();
    }

    @Override
    public BorderStyle getBorderLeft() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setBorderLeft(final BorderStyle border) {
        throw new UnsupportedOperationException();
    }

    @Override
    public BorderStyle getBorderLeftEnum() {
        throw new UnsupportedOperationException();
    }

    @Override
    public BorderStyle getBorderRight() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setBorderRight(final BorderStyle border) {
        throw new UnsupportedOperationException();
    }

    @Override
    public BorderStyle getBorderRightEnum() {
        throw new UnsupportedOperationException();
    }

    @Override
    public BorderStyle getBorderTop() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setBorderTop(final BorderStyle border) {
        throw new UnsupportedOperationException();
    }

    @Override
    public BorderStyle getBorderTopEnum() {
        throw new UnsupportedOperationException();
    }

    @Override
    public short getBottomBorderColor() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setBottomBorderColor(final short color) {
        throw new UnsupportedOperationException();
    }

    @Override
    public short getFillBackgroundColor() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setFillBackgroundColor(final short bg) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Color getFillBackgroundColorColor() {
        throw new UnsupportedOperationException();
    }

    @Override
    public short getFillForegroundColor() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setFillForegroundColor(final short bg) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Color getFillForegroundColorColor() {
        throw new UnsupportedOperationException();
    }

    @Override
    public FillPatternType getFillPattern() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setFillPattern(final FillPatternType fp) {
        throw new UnsupportedOperationException();
    }

    @Override
    public FillPatternType getFillPatternEnum() {
        throw new UnsupportedOperationException();
    }

    @Override
    public short getFontIndex() {
        throw new UnsupportedOperationException();
    }

    @Override
    public int getFontIndexAsInt() {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean getHidden() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setHidden(final boolean hidden) {
        throw new UnsupportedOperationException();
    }

    @Override
    public short getIndention() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setIndention(final short indent) {
        throw new UnsupportedOperationException();
    }

    @Override
    public short getIndex() {
        throw new UnsupportedOperationException();
    }

    @Override
    public short getLeftBorderColor() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setLeftBorderColor(final short color) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean getLocked() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setLocked(final boolean locked) {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean getQuotePrefixed() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setQuotePrefixed(final boolean quotePrefix) {

    }

    @Override
    public short getRightBorderColor() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setRightBorderColor(final short color) {

    }

    @Override
    public short getRotation() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setRotation(final short rotation) {

    }

    @Override
    public boolean getShrinkToFit() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setShrinkToFit(final boolean shrinkToFit) {

    }

    @Override
    public short getTopBorderColor() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setTopBorderColor(final short color) {

    }

    @Override
    public VerticalAlignment getVerticalAlignment() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setVerticalAlignment(final VerticalAlignment align) {

    }

    @Override
    public VerticalAlignment getVerticalAlignmentEnum() {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean getWrapText() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setWrapText(final boolean wrapped) {

    }

    @Override
    public void setDataFormat(final short fmt) {

    }

    @Override
    public void setFont(final Font font) {

    }
}
