package com.gaoice.easyexcel.writer.handler;

import com.gaoice.easyexcel.writer.sheet.SheetBuilderContext;

/**
 * 只关注 字段值 的转换
 * 更加简洁的函数式接口
 *
 * @author gaoice
 */
@FunctionalInterface
public interface FieldValueConverter<V> extends FieldHandler<V> {

    @Override
    default void handle(SheetBuilderContext<V> context) {
        context.setConvertedValue(convert(context.getValue()));
    }

    Object convert(V value);
}
