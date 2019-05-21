package com.gaoice.easyexcel.style;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 对标题 列名 统计结果 字体加粗
 */
public class BoldSheetStyle extends DefaultSheetStyle {
    @Override
    public XSSFCellStyle getTitleCellStyle(XSSFWorkbook workbook) {
        if (titleCellStyle != null) {
            return titleCellStyle;
        }
        XSSFCellStyle style = super.getTitleCellStyle(workbook);
        style.getFont().setBold(true);
        return style;
    }

    @Override
    public XSSFCellStyle getColumnNamesCellStyle(XSSFWorkbook workbook, int index) {
        if (columnNamesCellStyle != null) {
            return columnNamesCellStyle;
        }
        XSSFCellStyle style = super.getColumnNamesCellStyle(workbook, index);
        style.getFont().setBold(true);
        return style;
    }

    @Override
    public XSSFCellStyle getColumnCountCellStyle(XSSFWorkbook workbook, int columnIndex, Object v) {
        if (columnCountCellStyle != null) {
            return columnCountCellStyle;
        }
        XSSFCellStyle style = super.getColumnCountCellStyle(workbook, columnIndex, v);
        style.getFont().setBold(true);
        return style;
    }
}
