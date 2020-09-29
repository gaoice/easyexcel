package com.gaoice.easyexcel.data;

import com.gaoice.easyexcel.SheetInfo;

/**
 * 单元格值转换接口
 *
 * @param <P> 单元格值类型
 * @param <R> 统计结果类型
 * @author gaoice
 */
public interface Converter<P, R> {

    /**
     * 单元格值转换，如：将时间类型的变量转换为格式化后的字符串
     *
     * @param sheetInfo   当前 SheetInfo
     * @param value       当前单元格对应的值
     * @param listIndex   当前单元格对应的值在 list 中的 index
     * @param columnIndex 当前单元格对应的值所属列的 index（从 0 开始）
     * @return 转换结果
     */
    R convert(final SheetInfo sheetInfo, P value, int listIndex, int columnIndex);
}
