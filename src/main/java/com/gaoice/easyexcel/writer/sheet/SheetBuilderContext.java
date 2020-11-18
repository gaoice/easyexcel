package com.gaoice.easyexcel.writer.sheet;

import com.gaoice.easyexcel.writer.SheetInfo;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * SheetBuilder 的上下文
 *
 * @author gaoice
 */
public interface SheetBuilderContext<V> {

    /**
     * 正在处理的字段值，基本类型会被转为对应的包装类型，如 int 对应 Integer
     *
     * @return value
     */
    V getValue();

    /**
     * 获取转换后的值
     *
     * @return convertedValue
     */
    Object getConvertedValue();

    /**
     * 当前正在处理的行号
     *
     * @return rowNum 范围从 0 开始
     */
    int getRowNum();

    /**
     * 当前字段值所属对象在 {@link SheetInfo} 的 list 集合中的索引
     *
     * @return listIndex 范围从 0 开始
     */
    int getListIndex();

    /**
     * 当前字段值对应的列在 {@link SheetInfo} 中 columnNames 数组中的索引
     *
     * @return columnIndex 范围从 0 开始
     */
    int getColumnIndex();

    /**
     * 当前正在使用的 SheetInfo
     *
     * @return SheetInfo
     */
    SheetInfo getSheetInfo();

    /**
     * 当前正在处理的单元格
     *
     * @return cell
     */
    SXSSFCell getCell();

    /**
     * 当前正在处理的行
     *
     * @return row
     */
    SXSSFRow getRow();

    /**
     * 当前正在使用的 workbook
     *
     * @return workbook
     */
    SXSSFWorkbook getWorkbook();

    /**
     * 获取统计结果
     */
    Object getCountedResult();

    /**
     * 获取当前所有列的最大的字节长度，作为列宽
     */
    int[] getColumnMaxByteLengths();

    /**
     * 设置转换后的值
     */
    void setConvertedValue(Object o);

    /**
     * 设置统计结果
     */
    void setCountedResult(Object r);

    /**
     * 是否已经设置了单元格的样式
     */
    void setCellStyleHandled(boolean cellStyleHandled);

    /**
     * 是否已经设置了单元格的值
     */
    void setCellValueHandled(boolean cellValueHandled);

    /**
     * 是否已经设置了单元格的列宽（字节长度）
     */
    void setColumnByteLengthHandled(boolean columnLengthHandled);
}
