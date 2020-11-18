package com.gaoice.easyexcel.reader.converter;

import com.gaoice.easyexcel.reader.sheet.SheetParserContext;

/**
 * 解析 excel 时，单元格到类的字段的转换器
 *
 * @author gaoice
 */
@FunctionalInterface
public interface CellConverter {

    /**
     * @param context 上下文
     * @return 转换后的值，类型和要赋值的 Field 的类型一致
     */
    Object convert(SheetParserContext context);
}
