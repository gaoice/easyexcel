package com.gaoice.easyexcel.sheet;

import com.gaoice.easyexcel.SheetInfo;
import com.gaoice.easyexcel.data.Converter;
import com.gaoice.easyexcel.data.Counter;
import com.gaoice.easyexcel.style.DefaultSheetStyle;
import com.gaoice.easyexcel.style.SheetStyle;
import com.gaoice.easyexcel.util.Assert;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.UnsupportedEncodingException;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

/**
 * @author gaoice
 * @since 1.1
 */
public class SheetBuilder {
    private static final char VIRTUAL_FIELD_SIGN = '#';
    private static final String CHAR_SET = "utf-8";

    private SXSSFWorkbook workbook;
    private SheetInfo sheetInfo;

    private String sheetName;
    private String title;
    private String[] columnNames;
    private String[] fieldNames;
    private List<?> list;
    private SheetStyle sheetStyle;
    private Map<String, Object> converterMap;
    private Map<String, Object> counterMap;

    private List<List<String>> fieldNameBunches;
    private List<List<Field>> fieldBunches;
    private int columnCount;
    private int[] columnMaxLength;
    private Object[] countResults;

    private int nextRowNum = 0;

    private SXSSFSheet sheet;

    /**
     * 当前正在处理的 row
     */
    private SXSSFRow row;
    /**
     * 当前正在处理的 cell
     */
    private SXSSFCell cell;

    /**
     * 当前正在处理的 list 索引，在 list 的循环处理中使用
     */
    private int listIndex = 0;
    /**
     * 当前正在处理的 列 索引，在 list 的循环处理中使用
     */
    private int columnIndex = 0;

    /**
     * 此方法会返回 sheet，链式调用不会返回 sheet
     */
    public SXSSFSheet buildSheet(SXSSFWorkbook workbook, SheetInfo sheetInfo) throws UnsupportedEncodingException, NoSuchFieldException, IllegalAccessException {
        workbook(workbook).sheetInfo(sheetInfo).build();
        return sheet;
    }


    public static SheetBuilder builder() {
        return new SheetBuilder();
    }

    public SheetBuilder workbook(SXSSFWorkbook workbook) {
        Assert.notNull(workbook, "SXSSFWorkbook must be non-null");
        this.workbook = workbook;
        return this;
    }

    /**
     * 处理 SheetInfo
     */
    public SheetBuilder sheetInfo(SheetInfo sheetInfo) {
        Assert.notNull(sheetInfo, "SheetInfo must be non-null");
        this.sheetInfo = sheetInfo;
        sheetName = sheetInfo.getSheetName();
        title = sheetInfo.getTitle();
        columnNames = sheetInfo.getColumnNames();
        Assert.notEmpty(columnNames, sheetName + " columnNames must be non-null");
        fieldNames = sheetInfo.getFieldNames();
        Assert.notEmpty(fieldNames, sheetName + " fieldNames must be non-null");
        if (columnNames.length != fieldNames.length) {
            throw new IllegalArgumentException(sheetName + " columnNames.length shall equals fieldNames.length");
        }
        list = sheetInfo.getList();
        sheetStyle = sheetInfo.getSheetStyle();
        if (sheetStyle == null) {
            sheetStyle = new DefaultSheetStyle();
        }
        converterMap = sheetInfo.getConverterMap();
        counterMap = sheetInfo.getCounterMap();

        fieldNameBunches = new ArrayList<>(fieldNames.length);
        for (String classFieldName : fieldNames) {
            fieldNameBunches.add(Arrays.asList(classFieldName.split("\\.")));
        }
        fieldBunches = sheetInfo.getFieldCache();

        columnCount = fieldNames.length;
        columnMaxLength = new int[columnCount];
        countResults = counterMap == null ? null : new Object[columnCount];
        sheet = workbook.createSheet(sheetName);

        return this;
    }

    public void build() throws UnsupportedEncodingException, NoSuchFieldException, IllegalAccessException {
        buildTitle();
        buildColumnNames();
        buildList();
        buildCountResults();
        buildColumnWidth();
        finishSheetStyle();
        finishSheetInfo();
    }

    /**
     * 处理 title
     */
    private void buildTitle() {
        if (title != null && !title.isEmpty()) {
            sheet.addMergedRegion(new CellRangeAddress(nextRowNum, nextRowNum, 0, columnCount - 1));
            row = sheet.createRow(nextRowNum++);
            cell = row.createCell(0);
            cell.setCellValue(title);
            cell.setCellStyle(sheetStyle.getTitleCellStyle(workbook));
        }
    }

    /**
     * 处理列名
     */
    private void buildColumnNames() throws UnsupportedEncodingException {
        row = sheet.createRow(nextRowNum++);
        for (int i = 0; i < columnCount; i++) {
            cell = row.createCell(i);
            cell.setCellValue(columnNames[i]);
            cell.setCellStyle(sheetStyle.getColumnNamesCellStyle(workbook, i));
            int valueLength = columnNames[i].getBytes(CHAR_SET).length;
            if (valueLength > columnMaxLength[i]) {
                columnMaxLength[i] = valueLength;
            }
        }
    }

    /**
     * 处理 list
     */
    private void buildList() throws NoSuchFieldException, IllegalAccessException, UnsupportedEncodingException {
        if (list != null && list.size() > 0) {
            Object o = list.get(0);
            if (o instanceof Map) {
                processMapList();
            } else {
                processBeanList();
            }
        }
    }

    /**
     * 处理统计结果
     */
    private void buildCountResults() throws UnsupportedEncodingException {
        if (countResults != null) {
            row = sheet.createRow(nextRowNum++);
            for (int i = 0; i < columnCount; i++) {
                String value = countResults[i] == null ? "" : countResults[i].toString();
                cell = row.createCell(i);
                cell.setCellValue(value);
                cell.setCellStyle(sheetStyle.getColumnCountCellStyle(workbook, i, countResults[i]));
                int valueLength = value.getBytes(CHAR_SET).length;
                if (valueLength > columnMaxLength[i]) {
                    columnMaxLength[i] = valueLength;
                }
            }
        }
    }

    /**
     * 自适应列宽
     */
    private void buildColumnWidth() {
        sheetStyle.handleColumnMaxBytesLength(columnMaxLength);
        for (int i = 0; i < columnCount; i++) {
            sheet.setColumnWidth(i, columnMaxLength[i] * 256);
        }
    }

    private void finishSheetStyle() {
        sheetStyle.clearStyleCache();
    }

    private void finishSheetInfo() {
        sheetInfo.setFieldCache(fieldBunches);
    }

    private void processBeanList() throws NoSuchFieldException, IllegalAccessException, UnsupportedEncodingException {
        parseBeanField();
        for (listIndex = 0; listIndex < list.size(); listIndex++) {
            row = sheet.createRow(nextRowNum++);
            Object baseObject = list.get(listIndex);
            for (columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                Object o = baseObject;
                for (Field f : fieldBunches.get(columnIndex)) {
                    if (o != null) {
                        o = f.get(o);
                    } else {
                        break;
                    }
                }
                processListCellStyle(o);
                count(o);
                Object convertValue = convert(o);
                processListCellValue(convertValue);
            }
        }
    }

    private void processMapList() throws UnsupportedEncodingException {
        for (listIndex = 0; listIndex < list.size(); listIndex++) {
            row = sheet.createRow(nextRowNum++);
            Map<?, ?> map = (Map<?, ?>) list.get(listIndex);
            for (columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                Object o = map.get(fieldNames[columnIndex]);
                processListCellStyle(o);
                count(o);
                Object convertValue = convert(o);
                processListCellValue(convertValue);
            }
        }
    }

    /**
     * 对值进行转换
     *
     * @param object 待转换的值
     * @return 转换后的结果
     */
    private Object convert(Object object) {
        String fieldName = fieldNames[columnIndex];
        if (converterMap != null && converterMap.containsKey(fieldName)) {
            Object converter = converterMap.get(fieldName);
            if (converter instanceof Converter) {
                Converter<Object, Object> cellConverter = (Converter<Object, Object>) converter;
                object = cellConverter.convert(sheetInfo, object, listIndex, columnIndex);
            } else if (converter instanceof Map) {
                Map<?, ?> dictionary = (Map<?, ?>) converter;
                object = dictionary.get(object);
            }
        }
        return object;
    }

    /**
     *
     */
    private void count(Object object) {
        String fieldName = fieldNames[columnIndex];
        if (counterMap != null && counterMap.containsKey(fieldName)) {
            Object counter = counterMap.get(fieldName);
            if (counter instanceof Counter) {
                Counter<Object, Object> columnCounter = (Counter<Object, Object>) counter;
                countResults[columnIndex] = columnCounter.count(sheetInfo, object, listIndex, columnIndex, countResults[columnIndex]);
            } else if (object instanceof BigDecimal) {
                BigDecimal o1 = (BigDecimal) object;
                BigDecimal r = (BigDecimal) countResults[columnIndex];
                countResults[columnIndex] = r == null ? o1 : o1.add(r);
            } else if (object instanceof Number) {
                Number o1 = (Number) object;
                Number r = (Number) countResults[columnIndex];
                countResults[columnIndex] = r == null ? o1 : (o1.doubleValue() + r.doubleValue());
            } else {
                countResults[columnIndex] = "not support default counter";
            }
        }
    }

    /**
     * 处理单元格样式，参数使用的是转换前的值
     */
    private void processListCellStyle(Object object) {
        cell = row.createCell(columnIndex);
        cell.setCellStyle(sheetStyle.getListCellStyle(workbook, listIndex, columnIndex, object));
    }

    /**
     * 设置单元格的值，参数使用的是转换过后的值
     */
    private void processListCellValue(Object object) throws UnsupportedEncodingException {
        String value = object == null ? "" : object.toString();
        if (object instanceof Number) {
            Number num = (Number) object;
            cell.setCellValue(num.doubleValue());
        } else {
            cell.setCellValue(value);
        }
        int valueLength = value.getBytes(CHAR_SET).length;
        if (valueLength > columnMaxLength[columnIndex]) {
            columnMaxLength[columnIndex] = valueLength;
        }
    }

    private void parseBeanField() throws NoSuchFieldException {
        if (fieldBunches == null) {
            fieldBunches = new ArrayList<>(columnCount);
            Class<?> baseClass = list.get(0).getClass();
            for (int i = 0; i < columnCount; i++) {
                Class<?> c = baseClass;
                List<String> fieldNameBunch = fieldNameBunches.get(i);
                List<Field> fieldBunch = new ArrayList<>();
                String firstFieldName = fieldNameBunch.get(0);
                if (!"".equals(firstFieldName) && firstFieldName.charAt(0) != VIRTUAL_FIELD_SIGN) {
                    for (String fieldName : fieldNameBunch) {
                        Field f = c.getDeclaredField(fieldName);
                        f.setAccessible(true);
                        fieldBunch.add(f);
                        c = f.getType();
                    }
                }
                fieldBunches.add(fieldBunch);
            }
        }
    }
}
