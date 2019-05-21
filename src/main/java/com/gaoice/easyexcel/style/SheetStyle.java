package com.gaoice.easyexcel.style;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public interface SheetStyle {

    XSSFCellStyle getTitleCellStyle(XSSFWorkbook workbook);

    XSSFCellStyle getColumnNamesCellStyle(XSSFWorkbook workbook, int columnIndex);

    XSSFCellStyle getListCellStyle(XSSFWorkbook workbook, int listIndex, int columnIndex, Object v);

    XSSFCellStyle getColumnCountCellStyle(XSSFWorkbook workbook, int columnIndex, Object v);

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
