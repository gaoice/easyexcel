package com.gaoice.easyexcel.data;

import com.gaoice.easyexcel.SheetInfo;

/**
 * 列 计数/统计 接口
 *
 * @param <P> 单元格值类型
 * @param <R> 统计结果类型
 * @author gaoice
 */
public interface Counter<P, R> {

    /**
     * 对值进行统计，如：value 代表分数，则可以统计平均分
     *
     * @param sheetInfo   当前 SheetInfo
     * @param value       当前单元格对应的值
     * @param listIndex   当前单元格对应的值在 list 中的 index
     * @param columnIndex 当前单元格对应的值所属列的 index（从 0 开始）
     * @param result      此列的上一次统计结果
     * @return 本次的列统计结果，将作为下一次统计调用的传参
     */
    R count(final SheetInfo sheetInfo, P value, int listIndex, int columnIndex, R result);
}
