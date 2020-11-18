package com.gaoice.easyexcel.writer.handler;

import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;

/**
 * 一些默认的转换器实现
 *
 * @author gaoice
 */
public interface FieldHandlers {

    FieldValueConverter<Date> DATE_FORMATTER = value ->
            value == null ? null : new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(value);

    FieldValueConverter<LocalDateTime> LOCAL_DATE_TIME_FORMATTER = value ->
            value == null ? null : DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss").format(value);

    FieldValueConverter<LocalDate> LOCAL_DATE_FORMATTER = value ->
            value == null ? null : DateTimeFormatter.ofPattern("yyyy-MM-dd").format(value);

    FieldValueConverter<LocalTime> LOCAL_TIME_FORMATTER = value ->
            value == null ? null : DateTimeFormatter.ofPattern("HH:mm:ss").format(value);

    /**
     * Number 类型的字段写入 excel 为数字类型，使用此转换器可以转换为字符串类型
     */
    FieldValueConverter<Object> TO_STRING = value -> value == null ? null : value.toString();

    /**
     * order列转换和统计功能
     */
    FieldHandler<Object> ORDER_GENERATOR_RESULT_TEXT = context -> {
        context.setConvertedValue(context.getListIndex() + 1);
        context.setCountedResult("合计");
    };

    FieldHandler<Object> ORDER_GENERATOR = context ->
            context.setConvertedValue(context.getListIndex() + 1);

    FieldValueCounter<Object> ORDER_RESULT_TEXT = context -> "合计";
}
