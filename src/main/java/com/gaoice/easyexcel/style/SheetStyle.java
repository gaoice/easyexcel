package com.gaoice.easyexcel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * 一个 workbook 能够创建的 CellStyle 数量有限制
 * 所以对于每个返回 CellStyle 的方法，在第二次调用该方法的时候应该尽量使用缓存的 CellStyle
 *
 * @author gaoice
 */
public interface SheetStyle {

    /**
     * 获取 标题 样式
     *
     * @param workbook CellStyle 需要调用 workbook.createCellStyle() 创建出来
     * @return cell style
     */
    CellStyle getTitleCellStyle(SXSSFWorkbook workbook);

    /**
     * 获取 列名字（表头） 样式
     *
     * @param workbook    CellStyle 需要调用 workbook.createCellStyle() 创建出来
     * @param columnIndex 列索引序号，可以通过序号判断，对特殊的列使用特殊的样式
     * @return cell style
     */
    CellStyle getColumnNamesCellStyle(SXSSFWorkbook workbook, int columnIndex);

    /**
     * 获取 表格 样式
     *
     * @param workbook    CellStyle 需要调用 workbook.createCellStyle() 创建出来
     *                    由于 list 对应的单元格可能会较多，这个方法会调用多次，相同的样式应该缓存起来重复使用
     * @param listIndex   单元格值位于 list 中的索引，可用于辅助判断使用哪个 CellStyle ，如：判断奇数行还是偶数行，组成颜色深浅相间的表格
     * @param columnIndex 单元格值位于哪一列
     * @param v           单元格的值，不要在这个方法里对值进行写操作，要操作单元格的值使用 Converter
     * @return cell style
     */
    CellStyle getListCellStyle(SXSSFWorkbook workbook, int listIndex, int columnIndex, final Object v);

    /**
     * 获取 统计结果行 样式
     *
     * @param workbook    CellStyle 需要调用 workbook.createCellStyle() 创建出来
     * @param columnIndex 单元格值位于哪一列
     * @param v           单元格的值
     * @return cell style
     */
    CellStyle getColumnCountCellStyle(SXSSFWorkbook workbook, int columnIndex, Object v);

    /**
     * 列最大字节长度处理，用于单元格宽度自适应
     *
     * @param columnMaxBytesLength 每列的最大字节长度
     */
    default void handleColumnMaxBytesLength(int[] columnMaxBytesLength) {
    }

    /**
     * 列最大字节长度处理，用于单元格宽度自适应
     *
     * @param columnMaxBytesLength 每列的最大字节长度
     * @see #handleColumnMaxBytesLength(int[])
     */
    @Deprecated
    default void columnMaxBytesLengthHandler(int[] columnMaxBytesLength) {
        handleColumnMaxBytesLength(columnMaxBytesLength);
    }

    /**
     * 不同workbook style不能通用
     * ExcelBuilder在工作完成后将会调用此方法清除style缓存
     * 没有重复使用 SheetStyle 不必实现此方法
     */
    default void clearStyleCache() {
    }
}
