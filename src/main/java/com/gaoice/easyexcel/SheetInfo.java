package com.gaoice.easyexcel;

import com.gaoice.easyexcel.data.Converter;
import com.gaoice.easyexcel.data.Counter;
import com.gaoice.easyexcel.style.SheetStyle;
import com.gaoice.easyexcel.util.Assert;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author gaoice
 */
public class SheetInfo {
    private String sheetName;
    private String title;
    /**
     * 列名
     */
    private String[] columnNames;
    /**
     * 列名对应的类字段名
     */
    private String[] fieldNames;

    private List<?> list;
    private SheetStyle sheetStyle;
    private Map<String, Object> converterMap;
    private Map<String, Object> counterMap;
    /**
     * SheetInfo 被使用后，fieldCache 会缓存，相同的 fieldNames 可以重复使用 SheetInfo 以提高效率
     */
    private List<List<Field>> fieldCache;

    public SheetInfo() {
    }

    public SheetInfo(String sheetName, String[] columnNames, String[] fieldNames, List<?> list) {
        this.sheetName = sheetName;
        this.columnNames = columnNames;
        this.fieldNames = fieldNames;
        this.list = list;
    }

    public SheetInfo(String sheetName, String title, String[] columnNames, String[] fieldNames, List<?> list) {
        this.sheetName = sheetName;
        this.title = title;
        this.columnNames = columnNames;
        this.fieldNames = fieldNames;
        this.list = list;
    }

    /**
     * 为 fieldName列 设置单元格转换器，Map键值对
     *
     * @param fieldName  字段名
     * @param dictionary 数据字典
     */
    public SheetInfo putConverter(String fieldName, Map<?, ?> dictionary) {
        Assert.notNull(fieldName, "fieldName must be non-null");
        Assert.notNull(dictionary, "dictionary must be non-null");
        if (converterMap == null) {
            converterMap = new HashMap<>();
        }
        if (isAtFieldNames(fieldName)) {
            converterMap.put(fieldName, dictionary);
        } else {
            throw new IllegalArgumentException("no such fieldName, please put in array fieldNames");
        }
        return this;
    }

    /**
     * 为 fieldName列 设置单元格转换器，函数式接口
     *
     * @param fieldName 字段名
     * @param converter 转换器
     */
    public SheetInfo putConverter(String fieldName, Converter<?, ?> converter) {
        Assert.notNull(fieldName, "fieldName must be non-null");
        Assert.notNull(converter, "converter must be non-null");
        if (converterMap == null) {
            converterMap = new HashMap<>();
        }
        if (isAtFieldNames(fieldName)) {
            converterMap.put(fieldName, converter);
        } else {
            throw new IllegalArgumentException("no such fieldName, please put in array fieldNames");
        }
        return this;
    }

    /**
     * 为 fieldName列 设置列计数器，函数式接口
     *
     * @param fieldName 字段名
     * @param counter   计数器
     */
    public SheetInfo putCounter(String fieldName, Counter<?, ?> counter) {
        Assert.notNull(fieldName, "fieldName must be non-null");
        if (counterMap == null) {
            counterMap = new HashMap<>();
        }
        if (isAtFieldNames(fieldName)) {
            counterMap.put(fieldName, counter);
        } else {
            throw new IllegalArgumentException("no such fieldName, please put in array fieldNames");
        }
        return this;
    }

    private boolean isAtFieldNames(String fieldName) {
        boolean b = false;
        if (fieldNames != null) {
            for (String name : fieldNames) {
                if (name.equals(fieldName)) {
                    b = true;
                    break;
                }
            }
        }
        return b;
    }

    public String getSheetName() {
        return sheetName;
    }

    public SheetInfo setSheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    public String getTitle() {
        return title;
    }

    public SheetInfo setTitle(String title) {
        this.title = title;
        return this;
    }

    public String[] getColumnNames() {
        return columnNames;
    }

    public SheetInfo setColumnNames(String[] columnNames) {
        this.columnNames = columnNames;
        return this;
    }

    public String[] getFieldNames() {
        return fieldNames;
    }

    public SheetInfo setFieldNames(String[] fieldNames) {
        this.fieldNames = fieldNames;
        return this;
    }

    /**
     * @see #getFieldNames
     */
    @Deprecated
    public String[] getClassFieldNames() {
        return fieldNames;
    }

    /**
     * @see #setFieldNames(String[])
     */
    @Deprecated
    public void setClassFieldNames(String[] fieldNames) {
        this.fieldNames = fieldNames;
    }

    public List<?> getList() {
        return list;
    }

    public SheetInfo setList(List<?> list) {
        this.list = list;
        return this;
    }

    public SheetStyle getSheetStyle() {
        return sheetStyle;
    }

    public SheetInfo setSheetStyle(SheetStyle sheetStyle) {
        this.sheetStyle = sheetStyle;
        return this;
    }

    public Map<String, Object> getConverterMap() {
        return converterMap;
    }

    public Map<String, Object> getCounterMap() {
        return counterMap;
    }

    public List<List<Field>> getFieldCache() {
        return fieldCache;
    }

    public void setFieldCache(List<List<Field>> fieldCache) {
        this.fieldCache = fieldCache;
    }

    public SXSSFWorkbook build() throws Exception {
        return ExcelBuilder.createWorkbook(this);
    }

    public void build(OutputStream out) throws Exception {
        ExcelBuilder.writeOutputStream(this, out);
    }
}
