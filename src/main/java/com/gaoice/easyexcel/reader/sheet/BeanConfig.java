package com.gaoice.easyexcel.reader.sheet;

import com.gaoice.easyexcel.reader.converter.CellConverter;
import com.gaoice.easyexcel.util.Assert;

import java.util.HashMap;
import java.util.Map;

/**
 * @author gaoice
 */
public class BeanConfig<T> extends SheetConfig {

    /**
     * List 元素类型
     */
    private Class<T> targetClass;

    /**
     * 字段名，映射为实体类对应的字段名，支持多层表示，如：a.b.c，要忽略的列对应的字段可以设置为 ""
     */
    private String[] fieldNames;

    /**
     * 列转换器，如将 excel 中读取的日期转换为 Date 字段类型
     */
    private final Map<String, CellConverter> converterMap = new HashMap<>();

    /**
     * 为 fieldName 设置单元格转换器
     *
     * @param fieldName 字段名
     * @param converter 转换器
     */
    public BeanConfig<T> putConverter(String fieldName, CellConverter converter) {
        Assert.notNull(fieldName, "fieldName must be non-null");
        Assert.notNull(converter, "converter must be non-null");
        if (isAtFieldNames(fieldName)) {
            converterMap.put(fieldName, converter);
        } else {
            throw new IllegalArgumentException("no such fieldName, please put in array fieldNames");
        }
        return this;
    }

    private boolean isAtFieldNames(String fieldName) {
        if (fieldNames != null) {
            for (String name : fieldNames) {
                if (name.equals(fieldName)) {
                    return true;
                }
            }
        }
        return false;
    }

    public Class<T> getTargetClass() {
        return targetClass;
    }

    public BeanConfig<T> setTargetClass(Class<T> targetClass) {
        this.targetClass = targetClass;
        return this;
    }

    public String[] getFieldNames() {
        return fieldNames;
    }

    public BeanConfig<T> setFieldNames(String[] fieldNames) {
        this.fieldNames = fieldNames;
        return this;
    }

    public Map<String, CellConverter> getConverterMap() {
        return converterMap;
    }

    @Override
    public BeanConfig<T> setIgnoreException(boolean ignoreException) {
        super.setIgnoreException(ignoreException);
        return this;
    }

    @Override
    public BeanConfig<T> setListFirstRowNum(int listFirstRowNum) {
        super.setListFirstRowNum(listFirstRowNum);
        return this;
    }

    @Override
    public BeanConfig<T> setListLastRowNum(int listLastRowNum) {
        super.setListLastRowNum(listLastRowNum);
        return this;
    }
}
