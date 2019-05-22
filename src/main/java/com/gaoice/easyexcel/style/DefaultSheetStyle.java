package com.gaoice.easyexcel.style;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class DefaultSheetStyle implements SheetStyle {

    /**
     * 对style缓存
     * Workbook中创建单元格样式个数是有限制的
     */
    protected CellStyle titleCellStyle;
    protected CellStyle columnNamesCellStyle;
    protected CellStyle listCellStyle;
    protected CellStyle columnCountCellStyle;

    @Override
    public CellStyle getTitleCellStyle(SXSSFWorkbook workbook) {
        if (titleCellStyle != null) {
            return titleCellStyle;
        }
        titleCellStyle = newSimpleStyle(workbook);
        titleCellStyle.setWrapText(false);
        Font font = workbook.getFontAt(titleCellStyle.getFontIndexAsInt());
        font.setFontHeightInPoints((short) 22);
        return titleCellStyle;
    }

    @Override
    public CellStyle getColumnNamesCellStyle(SXSSFWorkbook workbook, int index) {
        if (columnNamesCellStyle != null) {
            return columnNamesCellStyle;
        }
        columnNamesCellStyle = newSimpleStyle(workbook);
        return columnNamesCellStyle;
    }

    @Override
    public CellStyle getListCellStyle(SXSSFWorkbook workbook, int listIndex, int columnIndex, Object v) {
        if (listCellStyle != null) {
            return listCellStyle;
        }
        listCellStyle = newSimpleStyle(workbook);
        return listCellStyle;
    }

    @Override
    public CellStyle getColumnCountCellStyle(SXSSFWorkbook workbook, int columnIndex, Object v) {
        if (columnCountCellStyle != null) {
            return columnCountCellStyle;
        }
        columnCountCellStyle = newSimpleStyle(workbook);
        return columnCountCellStyle;
    }

    /**
     * 最大值255
     *
     * @param columnMaxBytesLength
     */
    @Override
    public void columnMaxBytesLengthHandler(int[] columnMaxBytesLength) {
        for (int i = 0; i < columnMaxBytesLength.length; i++) {
            if (columnMaxBytesLength[i] > 255) {
                columnMaxBytesLength[i] = 255;
            }
        }
    }

    @Override
    public void clearStyleCache() {
        titleCellStyle = null;
        columnNamesCellStyle = null;
        listCellStyle = null;
        columnCountCellStyle = null;
    }

    public CellStyle newSimpleStyle(SXSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 12);
        style.setFont(font);
        return style;
    }
}
