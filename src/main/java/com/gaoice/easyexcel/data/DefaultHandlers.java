package com.gaoice.easyexcel.data;

import java.math.BigDecimal;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 提供一些 转换器/统计器 实现
 *
 * @author gaoice
 */
public class DefaultHandlers {

    /**
     * 序号转换器，返回listIndex+1作为序号
     */
    public static Converter<Object, Object> orderConverter = (sheetInfo, value, listIndex, columnIndex) -> listIndex + 1;
    /**
     * 序号统计器，统计信息行显示“合计”字样
     */
    public static Counter<Object, Object> orderCounter = (sheetInfo, value, listIndex, columnIndex, result) -> result != null ? result : "合计";

    /**
     * BigDecimal类求和统计
     */
    public static Counter<Object, Object> bigDecimalCounter = (sheetInfo, value, listIndex, columnIndex, result) -> {
        if (value instanceof BigDecimal) {
            BigDecimal bdValue = (BigDecimal) value;
            BigDecimal bdResult = (BigDecimal) result;
            return bdResult == null ? bdValue : bdValue.add(bdResult);
        }
        return result;
    };
    /**
     * Number类求和统计
     */
    public static Counter<Object, Object> numberCounter = (sheetInfo, value, listIndex, columnIndex, result) -> {
        if (value instanceof Number) {
            Number numValue = (Number) value;
            Number numResult = (Number) result;
            return numResult == null ? numValue : (numValue.doubleValue() + numResult.doubleValue());
        }
        return result;
    };

    /**
     * 时间格式化 yyyy-MM-dd
     */
    public static Converter<Object, Object> dateConverter = (sheetInfo, o, listIndex, columnIndex) -> {
        if (o instanceof Date) {
            SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
            return format.format((Date) o);
        }
        return o;
    };
    /**
     * 剔除时间戳小数点
     */
    public static Converter<Object, Object> timestampConverter = (sheetInfo, o, i, j) -> {
        if (o instanceof Timestamp) {
            String s = o.toString();
            int index;
            char dot = '.';
            if ((index = s.indexOf(dot)) > -1) {
                s = s.substring(0, index);
            }
            return s;
        }
        return o;
    };
}
