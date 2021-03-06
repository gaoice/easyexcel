package com.gaoice.easyexcel.reader.converter;

import org.apache.commons.lang3.StringUtils;

/**
 * 当字段没有单独配置 CellToFieldConverter 时，假如类型不匹配，会根据字段类型查找注册的 CellToFieldConverter
 *
 * @author gaoice
 */
public abstract class CellConverters {

    public static CellConverter INTEGER = context -> {
        String value = context.getStringValue();
        if (StringUtils.isBlank(value)) {
            return null;
        }
        //使用 double 可以处理 1.0 这种带小数的情况
        return Double.valueOf(value).intValue();
    };

    public static CellConverter DOUBLE = context -> {
        String value = context.getStringValue();
        if (StringUtils.isBlank(value)) {
            return null;
        }
        return Double.valueOf(value);
    };

    public static CellConverter FLOAT = context -> {
        String value = context.getStringValue();
        if (StringUtils.isBlank(value)) {
            return null;
        }
        return Double.valueOf(value).floatValue();
    };

    public static CellConverter LONG = context -> {
        String value = context.getStringValue();
        if (StringUtils.isBlank(value)) {
            return null;
        }
        return Double.valueOf(value).longValue();
    };

    public static CellConverter SHORT = context -> {
        String value = context.getStringValue();
        if (StringUtils.isBlank(value)) {
            return null;
        }
        return Double.valueOf(value).shortValue();
    };

    public static CellConverter BOOLEAN = context -> {
        String value = context.getStringValue();
        if (StringUtils.isBlank(value)) {
            return null;
        }
        return Boolean.valueOf(value);
    };

    public static CellConverter BYTE = context -> {
        String value = context.getStringValue();
        if (StringUtils.isBlank(value)) {
            return null;
        }
        return Byte.valueOf(value);
    };

    public static CellConverter CHAR = context -> {
        String value = context.getStringValue();
        if (StringUtils.isBlank(value)) {
            return null;
        }
        return value.charAt(0);
    };
}
