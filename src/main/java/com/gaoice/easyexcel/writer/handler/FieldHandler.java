package com.gaoice.easyexcel.writer.handler;

import com.gaoice.easyexcel.writer.sheet.SheetBuilderContext;

/**
 * 生成 excel 时，类的字段处理器，在设置单元格的值、样式、列宽之前调用
 * <p>
 * 功能：
 * <p>
 * 对 字段值 进行转换，通过 {@link SheetBuilderContext#setConvertedValue(Object)} 设置转换后的值
 * <p>
 * 对 字段值 进行统计，通过 {@link SheetBuilderContext#setCountedResult(Object)} 设置统计结果
 * <p>
 * 通过 {@link SheetBuilderContext#getCell()} 获取单元格，自定义设置单元格的值、样式等，然后通过
 * {@link SheetBuilderContext#setCellValueHandled(boolean)} 和 {@link SheetBuilderContext#setCellStyleHandled(boolean)}
 * 设置为 true 避免被重复处理
 * <p>
 * 通过 {@link SheetBuilderContext#getColumnMaxByteLengths()} 获取每列的宽度，进行宽度的调整
 *
 * @param <V> Field 的类型
 * @author gaoice
 */
@FunctionalInterface
public interface FieldHandler<V> {

    /**
     * 字段值转换功能，如果只需要值转换功能可以使用 {@link FieldValueConverter}
     * {@link SheetBuilderContext#getValue()} 获取字段值，由字段值决定，可能为 null ，注意进行处理
     * {@link SheetBuilderContext#setConvertedValue(Object)} 设置转换后的值
     * <p>
     * 字段值统计功能，如果只需要值统计功能可以使用 {@link FieldValueCounter}
     * {@link SheetBuilderContext#getCountedResult()} 获取统计的结果，第一次一定为 null ，注意进行处理
     * {@link SheetBuilderContext#setCountedResult(Object)} 设置统计的结果
     * <p>
     * 其他功能
     * 通过 {@link SheetBuilderContext#getCell()} 等方法获取变量进行自定义操作
     *
     * @param context 不会为 null
     */
    void handle(SheetBuilderContext<V> context);
}
