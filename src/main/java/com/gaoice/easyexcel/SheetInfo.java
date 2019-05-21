package com.gaoice.easyexcel;

import com.gaoice.easyexcel.data.Converter;
import com.gaoice.easyexcel.data.Counter;
import com.gaoice.easyexcel.style.SheetStyle;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

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
    private String[] classFieldNames;

    private List list;
    private SheetStyle sheetStyle;
    private Map<String, Object> converterMap;
    private Map<String, Object> counterMap;
    /**
     * SheetInfo 被使用后，fieldCache 会缓存，相同的 classFieldNames 可以重复使用 SheetInfo 以提高效率
     */
    private List<List<Field>> fieldCache;

    public SheetInfo(String sheetName, String[] columnNames, String[] classFieldNames, List list) {
        this.sheetName = sheetName;
        this.columnNames = columnNames;
        this.classFieldNames = classFieldNames;
        this.list = list;
    }

    public SheetInfo(String sheetName, String title, String[] columnNames, String[] classFieldNames, List list) {
        this.sheetName = sheetName;
        this.title = title;
        this.columnNames = columnNames;
        this.classFieldNames = classFieldNames;
        this.list = list;
    }

    /**
     * 为 classFieldName列 设置单元格转换器，Map键值对
     *
     * @param classFieldName 字段名
     * @param dictionary     数据字典
     */
    public void putConverter(String classFieldName, Map<Object, Object> dictionary) {
        if (classFieldName == null || dictionary == null) {
            throw new IllegalArgumentException("param shall not be null");
        }
        if (converterMap == null) {
            converterMap = new HashMap<>();
        }
        if (isAtClassFieldNames(classFieldName)) {
            converterMap.put(classFieldName, dictionary);
        } else {
            throw new IllegalArgumentException("no such classFieldName, please put in array classFieldNames");
        }
    }

    /**
     * 为 classFieldName列 设置单元格转换器，函数式接口
     *
     * @param classFieldName 字段名
     * @param converter      转换器
     */
    public void putConverter(String classFieldName, Converter converter) {
        if (classFieldName == null || converter == null) {
            throw new IllegalArgumentException("param shall not be null");
        }
        if (converterMap == null) {
            converterMap = new HashMap<>();
        }
        if (isAtClassFieldNames(classFieldName)) {
            converterMap.put(classFieldName, converter);
        } else {
            throw new IllegalArgumentException("no such classFieldName, please put in array classFieldNames");
        }
    }

    /**
     * 为 classFieldName列 设置列计数器，函数式接口
     *
     * @param classFieldName 字段名
     * @param counter        计数器
     */
    public void putCounter(String classFieldName, Counter counter) {
        if (classFieldName == null) {
            throw new IllegalArgumentException("classFieldName shall not be null");
        }
        if (counterMap == null) {
            counterMap = new HashMap<>();
        }
        if (isAtClassFieldNames(classFieldName)) {
            counterMap.put(classFieldName, counter);
        } else {
            throw new IllegalArgumentException("no such classFieldName, please put in array classFieldNames");
        }
    }

    private boolean isAtClassFieldNames(String classFieldName) {
        boolean b = false;
        if (classFieldNames != null) {
            for (String name : classFieldNames) {
                if (name.equals(classFieldName)) {
                    b = true;
                    break;
                }
            }
        }
        return b;
    }

    protected List<List<Field>> getFieldCache() {
        return fieldCache;
    }

    protected void setFieldCache(List<List<Field>> fieldCache) {
        this.fieldCache = fieldCache;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String[] getColumnNames() {
        return columnNames;
    }

    public void setColumnNames(String[] columnNames) {
        this.columnNames = columnNames;
    }

    public String[] getClassFieldNames() {
        return classFieldNames;
    }

    public void setClassFieldNames(String[] classFieldNames) {
        this.classFieldNames = classFieldNames;
    }

    public List getList() {
        return list;
    }

    public void setList(List list) {
        this.list = list;
    }

    public SheetStyle getSheetStyle() {
        return sheetStyle;
    }

    public void setSheetStyle(SheetStyle sheetStyle) {
        this.sheetStyle = sheetStyle;
    }

    public Map<String, Object> getConverterMap() {
        return converterMap;
    }

    public Map<String, Object> getCounterMap() {
        return counterMap;
    }
}
