package com.gaoice.easyexcel.writer.handler;

import com.gaoice.easyexcel.writer.sheet.SheetBuilderContext;

/**
 * 只关注统计功能
 *
 * @author gaoice
 */
@FunctionalInterface
public interface FieldValueCounter<V> extends FieldHandler<V> {

    @Override
    default void handle(SheetBuilderContext<V> context) {
        context.setCountedResult(count(context));
    }

    Object count(SheetBuilderContext<V> context);
}
