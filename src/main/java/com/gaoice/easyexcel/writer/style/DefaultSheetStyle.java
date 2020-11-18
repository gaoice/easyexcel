package com.gaoice.easyexcel.writer.style;

import com.gaoice.easyexcel.writer.sheet.SheetBuilderContext;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * 默认的简单样式
 *
 * @author gaoice
 */
public class DefaultSheetStyle implements SheetStyle {

    /**
     * 对 style 缓存
     * Workbook 中创建单元格样式个数是有限制的
     */
    protected CellStyle titleCellStyle;
    protected CellStyle columnNamesCellStyle;
    protected CellStyle listCellStyle;
    protected CellStyle columnCountCellStyle;

    @Override
    public CellStyle getTitleCellStyle(SheetBuilderContext<?> context) {
        if (titleCellStyle != null) {
            return titleCellStyle;
        }
        titleCellStyle = newSimpleStyle(context);
        titleCellStyle.setWrapText(false);
        Font font = context.getWorkbook().getFontAt(titleCellStyle.getFontIndexAsInt());
        font.setFontHeightInPoints((short) 22);
        return titleCellStyle;
    }

    @Override
    public CellStyle getColumnNamesCellStyle(SheetBuilderContext<?> context) {
        if (columnNamesCellStyle != null) {
            return columnNamesCellStyle;
        }
        columnNamesCellStyle = newSimpleStyle(context);
        return columnNamesCellStyle;
    }

    @Override
    public CellStyle getListCellStyle(SheetBuilderContext<?> context) {
        if (listCellStyle != null) {
            return listCellStyle;
        }
        listCellStyle = newSimpleStyle(context);
        return listCellStyle;
    }

    @Override
    public CellStyle getColumnCountCellStyle(SheetBuilderContext<?> context) {
        if (columnCountCellStyle != null) {
            return columnCountCellStyle;
        }
        columnCountCellStyle = newSimpleStyle(context);
        return columnCountCellStyle;
    }

    /**
     * 可以自定义设置每列单元格最长的字节长度
     */
    @Override
    public void handleColumnMaxBytesLength(SheetBuilderContext<?> context) {
    }

    @Override
    public void clearStyleCache() {
        titleCellStyle = null;
        columnNamesCellStyle = null;
        listCellStyle = null;
        columnCountCellStyle = null;
    }

    public CellStyle newSimpleStyle(SheetBuilderContext<?> context) {
        return newSimpleStyle(context.getWorkbook());
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
