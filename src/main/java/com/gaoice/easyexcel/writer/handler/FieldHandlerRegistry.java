package com.gaoice.easyexcel.writer.handler;

import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.Date;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.ConcurrentHashMap;

/**
 * 处理器注册
 * 可以为某种类型的字段设置处理器，而不必每次新建 {@link com.gaoice.easyexcel.writer.SheetInfo} 时都进行设置
 * 例如，默认为 日期类型 设置了处理器，当字段是日期类型并且没有单独设置 FieldHandler 的时候，就会使用此处注册的默认处理器
 * 如果想要统一 yyyy-MM-dd HH:mm 格式的日期字符串
 * 使用 {@link FieldHandlerRegistry#register(Class, FieldHandler)} 覆盖原有的处理器即可
 *
 * @author gaoice
 */
public class FieldHandlerRegistry {
    private static final Map<Class<?>, FieldHandler<?>> FIELD_HANDLER_MAP = new ConcurrentHashMap<>();

    static {
        register(Date.class, FieldHandlers.DATE_FORMATTER);
        register(java.sql.Date.class, FieldHandlers.DATE_FORMATTER);
        register(Timestamp.class, FieldHandlers.DATE_FORMATTER);
        register(LocalDate.class, FieldHandlers.LOCAL_DATE_FORMATTER);
        register(LocalTime.class, FieldHandlers.LOCAL_TIME_FORMATTER);
        register(LocalDateTime.class, FieldHandlers.LOCAL_DATE_TIME_FORMATTER);
    }

    /**
     * @param clazz        字段类型
     * @param fieldHandler 字段处理器
     * @param <V>          字段类型
     */
    public static <V> void register(Class<? extends V> clazz, FieldHandler<V> fieldHandler) {
        FIELD_HANDLER_MAP.put(clazz, fieldHandler);
    }

    public static Optional<FieldHandler<?>> get(Class<?> clazz) {
        return Optional.ofNullable(FIELD_HANDLER_MAP.get(clazz));
    }

    public static void clear() {
        FIELD_HANDLER_MAP.clear();
    }

    public static void remove(Class<?> clazz) {
        FIELD_HANDLER_MAP.remove(clazz);
    }
}
