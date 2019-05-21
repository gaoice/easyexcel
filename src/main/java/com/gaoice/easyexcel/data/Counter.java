package com.gaoice.easyexcel.data;

import com.gaoice.easyexcel.SheetInfo;

/**
 * 列统计接口
 *
 * @param <P>
 * @param <R>
 */
public interface Counter<P, R> {
    /**
     * @param sheetInfo
     * @param value       由于 totalizer 在 converter 后被调用，value 是经过 converter 转换的值
     * @param listIndex   当前单元格内的值在 list 中的index
     * @param columnIndex 当前单元格内的值所属列的index（从 0 计数）
     * @param result      此列的列统计结果
     * @return
     */
    R count(final SheetInfo sheetInfo, P value, int listIndex, int columnIndex, R result);
}
