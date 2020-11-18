package com.gaoice.easyexcel.writer;

import com.gaoice.easyexcel.util.Assert;
import com.gaoice.easyexcel.writer.handler.FieldHandler;
import com.gaoice.easyexcel.writer.handler.FieldValueConverter;
import com.gaoice.easyexcel.writer.handler.FieldValueCounter;
import com.gaoice.easyexcel.writer.sheet.SheetBuilderContext;
import com.gaoice.easyexcel.writer.style.SheetStyle;
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

    private final Map<String, FieldHandler<?>> fieldHandlerMap = new HashMap<>();

    /**
     * SheetInfo 被使用后，fieldCache 会缓存，可以重复使用 SheetInfo 以提高效率
     */
    private List<List<Field>> fieldCache;

    public SheetInfo() {
    }

    public SheetInfo(String[] fieldNames, List<?> list) {
        this.sheetName = "default";
        this.columnNames = fieldNames;
        this.fieldNames = fieldNames;
        this.list = list;
    }

    public SheetInfo(String sheetName, String[] fieldNames, List<?> list) {
        this.sheetName = sheetName;
        this.columnNames = fieldNames;
        this.fieldNames = fieldNames;
        this.list = list;
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
     * 为 fieldName列 设置处理器
     *
     * @param fieldName    字段名
     * @param fieldHandler 处理器
     */
    public SheetInfo putFieldHandler(String fieldName, FieldHandler<?> fieldHandler) {
        Assert.notNull(fieldName, "fieldName must be non-null");
        Assert.notNull(fieldHandler, "fieldHandler must be non-null");
        if (isAtFieldNames(fieldName)) {
            fieldHandlerMap.put(fieldName, fieldHandler);
        } else {
            throw new IllegalArgumentException("no such fieldName, please put in array fieldNames");
        }
        return this;
    }

    public <V> SheetInfo putFieldHandler(String fieldName, FieldValueConverter<V> converter, FieldValueCounter<V> counter) {
        Assert.notNull(fieldName, "fieldName must be non-null");
        Assert.notNull(converter, "converter must be non-null");
        Assert.notNull(counter, "counter must be non-null");
        if (isAtFieldNames(fieldName)) {
            fieldHandlerMap.put(fieldName, new FieldHandlerAdapter<>(converter, counter));
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

    public Map<String, FieldHandler<?>> getFieldHandlerMap() {
        return fieldHandlerMap;
    }

    public List<List<Field>> getFieldCache() {
        return fieldCache;
    }

    public void setFieldCache(List<List<Field>> fieldCache) {
        this.fieldCache = fieldCache;
    }

    public SXSSFWorkbook build() throws Exception {
        return ExcelWriter.createWorkbook(this);
    }

    public void build(OutputStream out) throws Exception {
        ExcelWriter.writeOutputStream(this, out);
    }

    /**
     * FieldConverter FieldCounter
     */
    public static class FieldHandlerAdapter<V> implements FieldHandler<V> {
        private final FieldValueConverter<V> converter;
        private final FieldValueCounter<V> counter;

        public FieldHandlerAdapter(FieldValueConverter<V> converter, FieldValueCounter<V> counter) {
            this.converter = converter;
            this.counter = counter;
        }

        @Override
        public void handle(SheetBuilderContext<V> context) {
            context.setConvertedValue(converter.convert(context.getValue()));
            context.setCountedResult(counter.count(context));
        }
    }
}
