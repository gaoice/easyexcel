package com.gaoice.easyexcel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public interface SheetStyle {

    CellStyle getTitleCellStyle(SXSSFWorkbook workbook);

    CellStyle getColumnNamesCellStyle(SXSSFWorkbook workbook, int columnIndex);

    CellStyle getListCellStyle(SXSSFWorkbook workbook, int listIndex, int columnIndex, Object v);

    CellStyle getColumnCountCellStyle(SXSSFWorkbook workbook, int columnIndex, Object v);

    /**
     * 列最大字节长度处理
     *
     * @param columnMaxBytesLength
     * @return
     */
    void columnMaxBytesLengthHandler(int[] columnMaxBytesLength);

    /**
     * 不同workbook style不能通用
     * ExcelBuilder在工作完成后将会调用此方法清除style缓存
     */
    void clearStyleCache();
}
