package com.gaoice.easyexcel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public interface SheetStyle {

    CellStyle getTitleCellStyle(SXSSFWorkbook workbook);

    CellStyle getColumnNamesCellStyle(SXSSFWorkbook workbook, int columnIndex);

    CellStyle getListCellStyle(SXSSFWorkbook workbook, int listIndex, int columnIndex, Object v);

    CellStyle getColumnCountCellStyle(SXSSFWorkbook workbook, int columnIndex, Object v);

    /**
     * 列最大字节长度处理，用于单元格宽度自适应
     *
     * @param columnMaxBytesLength 每列的最大字节长度
     */
    void columnMaxBytesLengthHandler(int[] columnMaxBytesLength);

    /**
     * 不同workbook style不能通用
     * ExcelBuilder在工作完成后将会调用此方法清除style缓存
     * 没有重复使用 SheetStyle 不必实现此方法
     */
    default void clearStyleCache() {
    }
}
