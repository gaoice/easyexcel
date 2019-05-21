package com.gaoice.easyexcel.data;

import com.gaoice.easyexcel.SheetInfo;

/**
 * 单元格值转换接口
 *
 * @param <P>
 * @param <R>
 */
public interface Converter<P, R> {
    R convert(final SheetInfo sheetInfo, P value, int listIndex, int columnIndex);
}
