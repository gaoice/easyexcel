package com.gaoice.easyexcel.data;

import com.gaoice.easyexcel.SheetInfo;

import java.math.BigDecimal;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 提供一些 转换器/统计器 实现
 */
public class DefaultHandlers {

    /**
     * 序号转换器，返回listIndex+1作为序号
     */
    public static Converter orderConverter = new Converter() {
        public Object convert(SheetInfo sheetInfo, Object value, int listIndex, int columnIndex) {
            return listIndex + 1;
        }
    };
    /**
     * 序号统计器，统计信息行显示“合计”字样
     */
    public static Counter orderCounter = new Counter() {
        public Object count(SheetInfo sheetInfo, Object value, int listIndex, int columnIndex, Object result) {
            return result != null ? result : "合计";
        }
    };
    /**
     * BigDecimal类求和统计
     */
    public static Counter bigDecimalCounter = new Counter() {
        public Object count(SheetInfo sheetInfo, Object value, int listIndex, int columnIndex, Object result) {
            if (value instanceof BigDecimal) {
                BigDecimal bdValue = (BigDecimal) value;
                BigDecimal bdResult = (BigDecimal) result;
                return bdResult == null ? bdValue : bdValue.add(bdResult);
            }
            return result;
        }
    };
    /**
     * Number类求和统计
     */
    public static Counter numberCounter = new Counter() {
        public Object count(SheetInfo sheetInfo, Object value, int listIndex, int columnIndex, Object result) {
            if (value instanceof Number) {
                Number numValue = (Number) value;
                Number numResult = (Number) result;
                return numResult == null ? numValue : (numValue.doubleValue() + numResult.doubleValue());
            }
            return result;
        }
    };
    /**
     * 时间格式化 yyyy-MM-dd
     */
    public static Converter dateConverter = new Converter() {
        public Object convert(SheetInfo sheetInfo, Object o, int listIndex, int columnIndex) {
            if (o instanceof Date) {
                SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
                return format.format((Date) o);
            }
            return o;
        }
    };
    /**
     * 剔除时间戳小数点
     */
    public static Converter timestampConverter = new Converter() {
        public Object convert(SheetInfo sheetInfo, Object o, int i, int j) {
            if (o instanceof Timestamp) {
                String s = o.toString();
                int index;
                if ((index = s.indexOf(".")) > -1)
                    s = s.substring(0, index);
                return s;
            }
            return o;
        }
    };
}
