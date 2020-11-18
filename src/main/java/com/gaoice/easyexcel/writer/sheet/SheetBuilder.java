package com.gaoice.easyexcel.writer.sheet;

import com.gaoice.easyexcel.util.Assert;
import com.gaoice.easyexcel.util.ReflectionUtils;
import com.gaoice.easyexcel.writer.SheetInfo;
import com.gaoice.easyexcel.writer.handler.FieldHandler;
import com.gaoice.easyexcel.writer.handler.FieldHandlerRegistry;
import com.gaoice.easyexcel.writer.style.DefaultSheetStyle;
import com.gaoice.easyexcel.writer.style.SheetStyle;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.UnsupportedEncodingException;
import java.lang.reflect.Field;
import java.util.*;

/**
 * @author gaoice
 * @since 1.1
 */
public class SheetBuilder implements SheetBuilderContext<Object> {
    private static final char VIRTUAL_FIELD_SIGN = '#';
    private static final String CHAR_SET = "utf-8";

    private SXSSFWorkbook workbook;
    private SheetInfo sheetInfo;

    private String title;
    private String[] columnNames;
    private String[] fieldNames;
    private List<?> list;
    private SheetStyle sheetStyle;
    private Map<String, FieldHandler<?>> fieldHandlerMap;

    private List<List<String>> fieldNameChains;
    private List<List<Field>> fieldChains;
    private int columnCount;
    private int[] columnMaxByteLengths;
    private Object[] countedResults;
    private SXSSFSheet sheet;
    private Class<?> listElementClass;
    private boolean isMapList;

    /**
     * 当前正在处理的 行号
     */
    private int rowNum = 0;

    /**
     * 当前正在处理的 列索引
     */
    private int columnIndex = 0;

    /**
     * 当前正在处理的 row
     */
    private SXSSFRow row;

    /**
     * 当前正在处理的 cell
     */
    private SXSSFCell cell;

    /**
     * 当前正在处理的 list 索引，在处理 list 的循环中使用，-1 表示 list 未开始处理或者已经处理结束
     */
    private int listIndex = -1;

    /**
     * 当前正在处理的字段值，在处理 list 的循环处理中使用
     */
    private Object value;

    /**
     * 转换后的字段值，在处理 list 的循环处理中使用
     */
    private Object convertedValue;

    /**
     * 是否已经设置了单元格样式，{@link FieldHandler} 如果对 cell 进行了处理，可以设置为 true ，将不在重复处理
     */
    private boolean cellStyleHandled = false;

    /**
     * 是否已经设置了单元格的值，{@link FieldHandler} 如果对 cell 进行了处理，可以设置为 true ，将不在重复处理
     */
    private boolean cellValueHandled = false;

    /**
     * 是否已经设置了当前列的字节长度，{@link FieldHandler} 如果对 cell 进行了处理，可以设置为 true ，将不在重复处理
     */
    private boolean columnByteLengthHandled = false;

    /**
     * 此方法会返回 sheet
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
        Assert.notNull(sheetInfo, "sheetInfo must be non-null");
        this.sheetInfo = sheetInfo;
        String sheetName = sheetInfo.getSheetName();
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
        fieldHandlerMap = sheetInfo.getFieldHandlerMap();

        fieldNameChains = new ArrayList<>(fieldNames.length);
        for (String fieldName : fieldNames) {
            fieldNameChains.add(Arrays.asList(fieldName.split("\\.")));
        }
        fieldChains = sheetInfo.getFieldCache();

        columnCount = fieldNames.length;
        columnMaxByteLengths = new int[columnCount];
        sheet = workbook.createSheet(sheetName);
        listElementClass = getListElementClass();

        return this;
    }

    public void build() throws UnsupportedEncodingException, NoSuchFieldException, IllegalAccessException {
        buildTitle();
        buildColumnNames();
        buildList();
        buildCountedResults();
        buildColumnWidth();
        finishSheetStyle();
        finishSheetInfo();
    }

    /**
     * 处理 title
     */
    private void buildTitle() {
        if (title != null && !title.isEmpty()) {
            sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, 0, columnCount - 1));
            row = sheet.createRow(rowNum++);
            cell = row.createCell(0);
            cell.setCellValue(title);
            cell.setCellStyle(sheetStyle.getTitleCellStyle(this));
        }
    }

    /**
     * 处理列名
     */
    private void buildColumnNames() throws UnsupportedEncodingException {
        row = sheet.createRow(rowNum);
        for (columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            cell = row.createCell(columnIndex);
            cell.setCellValue(columnNames[columnIndex]);
            cell.setCellStyle(sheetStyle.getColumnNamesCellStyle(this));
            handleColumnMaxByteLength(columnNames[columnIndex].getBytes(CHAR_SET).length, columnIndex);
        }
        rowNum++;
    }

    /**
     * 处理 list
     */
    private void buildList() throws NoSuchFieldException, IllegalAccessException, UnsupportedEncodingException {
        if (list != null && list.size() > 0) {
            if (listElementClass == null) {
                return;
            }
            isMapList = Map.class.isAssignableFrom(listElementClass);
            if (isMapList) {
                processMapList();
                return;
            }
            processBeanList();
        }
    }

    /**
     * 处理统计结果
     */
    private void buildCountedResults() throws UnsupportedEncodingException {
        if (countedResults != null) {
            row = sheet.createRow(rowNum);
            for (columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                String value = countedResults[columnIndex] == null ? "" : countedResults[columnIndex].toString();
                cell = row.createCell(columnIndex);
                cell.setCellValue(value);
                cell.setCellStyle(sheetStyle.getColumnCountCellStyle(this));
                handleColumnMaxByteLength(value.getBytes(CHAR_SET).length, columnIndex);
            }
            rowNum++;
        }
    }

    /**
     * 自适应列宽
     */
    private void buildColumnWidth() {
        sheetStyle.handleColumnMaxBytesLength(this);
        for (int i = 0; i < columnCount; i++) {
            int l = columnMaxByteLengths[i];
            if (l > 255) {
                l = 255;
            }
            sheet.setColumnWidth(i, l * 256);
        }
    }

    private void finishSheetStyle() {
        sheetStyle.clearStyleCache();
    }

    private void finishSheetInfo() {
        sheetInfo.setFieldCache(fieldChains);
    }

    private void processBeanList() throws NoSuchFieldException, IllegalAccessException, UnsupportedEncodingException {
        parseBeanField();
        for (listIndex = 0; listIndex < list.size(); listIndex++) {
            row = sheet.createRow(rowNum);
            Object baseObject = list.get(listIndex);
            if (baseObject == null) {
                continue;
            }
            for (columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                if (fieldChains.get(columnIndex).size() != 0) {
                    value = ReflectionUtils.getFieldValue(baseObject, fieldChains.get(columnIndex));
                } else {
                    value = baseObject;
                }
                processListCell();
            }
            rowNum++;
        }
    }

    private void processMapList() throws UnsupportedEncodingException {
        for (listIndex = 0; listIndex < list.size(); listIndex++) {
            row = sheet.createRow(rowNum);
            Map<?, ?> map = (Map<?, ?>) list.get(listIndex);
            if (map == null) {
                continue;
            }
            for (columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                value = map.get(fieldNames[columnIndex]);
                processListCell();
            }
            rowNum++;
        }
    }

    private void processListCell() throws UnsupportedEncodingException {
        cell = row.createCell(columnIndex);
        convertedValue = value;
        resetHandledStatus();
        invokeHandler();
        if (!cellStyleHandled) {
            cell.setCellStyle(sheetStyle.getListCellStyle(this));
        }
        if (!cellValueHandled) {
            setListCellValue();
        }
        if (!columnByteLengthHandled && convertedValue != null) {
            handleColumnMaxByteLength(convertedValue.toString().getBytes(CHAR_SET).length, columnIndex);
        }
    }

    private void resetHandledStatus() {
        cellStyleHandled = false;
        cellValueHandled = false;
        columnByteLengthHandled = false;
    }

    /**
     * 对值进行转换
     */
    @SuppressWarnings("unchecked")
    private <V> void invokeHandler() {
        String fieldName = fieldNames[columnIndex];
        if (fieldHandlerMap.containsKey(fieldName)) {
            FieldHandler<V> handler = (FieldHandler<V>) fieldHandlerMap.get(fieldName);
            handler.handle((SheetBuilderContext<V>) this);
            return;
        }
        // 查询 FieldHandlerRegistry 中根据字段类型注册的默认处理器
        Class<?> targetClass;
        if (value != null) {
            targetClass = value.getClass();
        } else {
            if (isMapList) {
                return;
            }
            List<Field> fieldChain = fieldChains.get(columnIndex);
            int size = fieldChain.size();
            if (size == 0) {
                targetClass = listElementClass;
            } else {
                targetClass = fieldChain.get(size - 1).getType();
            }
        }
        Optional<FieldHandler<?>> optional = FieldHandlerRegistry.get(targetClass);
        if (optional.isPresent()) {
            FieldHandler<V> handler = (FieldHandler<V>) optional.get();
            handler.handle((SheetBuilderContext<V>) this);
        }
    }

    /**
     * 设置单元格的值，使用的是转换过后的值
     */
    private void setListCellValue() {
        if (convertedValue == null) {
            return;
        }
        if (convertedValue instanceof CharSequence) {
            cell.setCellValue(convertedValue.toString());
        } else if (convertedValue instanceof Number) {
            cell.setCellValue(((Number) convertedValue).doubleValue());
        } else if (convertedValue instanceof Date) {
            cell.setCellValue((Date) convertedValue);
        } else if (convertedValue instanceof Calendar) {
            cell.setCellValue((Calendar) convertedValue);
        } else if (convertedValue instanceof RichTextString) {
            cell.setCellValue((RichTextString) convertedValue);
        } else {
            cell.setCellValue(convertedValue.toString());
        }
    }

    private void handleColumnMaxByteLength(int byteLength, int columnIndex) {
        if (byteLength > columnMaxByteLengths[columnIndex]) {
            columnMaxByteLengths[columnIndex] = byteLength;
        }
    }

    private void parseBeanField() throws NoSuchFieldException {
        if (fieldChains == null) {
            fieldChains = new ArrayList<>(columnCount);
            for (int i = 0; i < columnCount; i++) {
                List<String> fieldNameChain = fieldNameChains.get(i);
                String firstFieldName = fieldNameChain.get(0);
                if (!"".equals(firstFieldName) && firstFieldName.charAt(0) != VIRTUAL_FIELD_SIGN) {
                    fieldChains.add(Arrays.asList(ReflectionUtils.getFieldChain(listElementClass, fieldNameChain)));
                } else {
                    fieldChains.add(Collections.emptyList());
                }
            }
        }
    }

    private Class<?> getListElementClass() {
        for (Object o : list) {
            if (o != null) {
                return o.getClass();
            }
        }
        return null;
    }

    @Override
    public Object getValue() {
        return value;
    }

    @Override
    public Object getConvertedValue() {
        return convertedValue;
    }

    @Override
    public int getRowNum() {
        return rowNum;
    }

    @Override
    public int getListIndex() {
        return listIndex;
    }

    @Override
    public int getColumnIndex() {
        return columnIndex;
    }

    @Override
    public SXSSFRow getRow() {
        return row;
    }

    @Override
    public SXSSFCell getCell() {
        return cell;
    }

    @Override
    public SXSSFWorkbook getWorkbook() {
        return workbook;
    }

    @Override
    public SheetInfo getSheetInfo() {
        return sheetInfo;
    }

    @Override
    public Object getCountedResult() {
        return countedResults == null ? null : countedResults[columnIndex];
    }

    @Override
    public int[] getColumnMaxByteLengths() {
        return columnMaxByteLengths;
    }

    @Override
    public void setCountedResult(Object r) {
        if (r == null) {
            return;
        }
        if (countedResults == null) {
            countedResults = new Object[columnCount];
        }
        countedResults[columnIndex] = r;
    }

    @Override
    public void setConvertedValue(Object v) {
        this.convertedValue = v;
    }

    @Override
    public void setCellStyleHandled(boolean cellStyleHandled) {
        this.cellStyleHandled = cellStyleHandled;
    }

    @Override
    public void setCellValueHandled(boolean cellValueHandled) {
        this.cellValueHandled = cellValueHandled;
    }

    @Override
    public void setColumnByteLengthHandled(boolean columnLengthHandled) {
        this.columnByteLengthHandled = columnLengthHandled;
    }
}
