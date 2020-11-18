package com.gaoice.easyexcel.reader.sheet;

import com.gaoice.easyexcel.exception.CellException;
import com.gaoice.easyexcel.reader.SheetResult;
import com.gaoice.easyexcel.reader.converter.CellConverter;
import com.gaoice.easyexcel.reader.converter.CellConverterRegistry;
import com.gaoice.easyexcel.util.Assert;
import com.gaoice.easyexcel.util.ReflectionUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.*;

/**
 * @author gaoice
 */
public class SheetParser implements SheetParserContext {

    private XSSFSheet sheet;
    private BeanConfig<?> beanConfig;

    /**
     * List 开始的行号
     */
    private int listFirstRowNum = -1;

    /**
     * List 结束的行号
     */
    private int listLastRowNum = Integer.MIN_VALUE;

    /**
     * 是否忽略单元格转换为字段值时候的异常
     */
    private boolean isIgnoreException = false;

    private Map<String, CellConverter> converterMap;

    private String[] fieldNames;
    private List<List<String>> fieldNameChains;
    private List<List<Field>> fieldChains;
    private Class<?> targetClass;

    private String[] columnNames;

    private final List<Map<Integer, Object>> mapList = new ArrayList<>();
    private final List<Object> beanList = new ArrayList<>();

    private final List<CellException> exceptions = new ArrayList<>();

    /**
     * 列数
     */
    private int columnCount;

    /**
     * 当前正在处理的行
     */
    private XSSFRow row;

    /**
     * 当前正在处理的单元格
     */
    private XSSFCell cell;

    /**
     * 当前正在处理的行号
     */
    private int rowNum = 0;

    /**
     * 当前正在处理的列索引
     */
    private int columnIndex = 0;

    /**
     * 当前正在处理的单元格的值
     */
    private Object value;
    private String stringValue;

    public static SheetParser parser() {
        return new SheetParser();
    }

    public SheetResult<Map<Integer, Object>> parseMapList(XSSFSheet sheet, SheetConfig config) {
        sheet(sheet);
        sheetConfig(config);
        parseMapList();
        return new SheetResult<>(mapList, columnNames, exceptions);
    }

    @SuppressWarnings("unchecked")
    public <T> SheetResult<T> parseBeanList(XSSFSheet sheet, BeanConfig<T> config) throws NoSuchFieldException, NoSuchMethodException, InstantiationException, IllegalAccessException, InvocationTargetException {
        sheet(sheet);
        beanConfig(config);
        parseBeanList();
        return new SheetResult<>((List<T>) beanList, columnNames, exceptions);
    }

    private void sheet(XSSFSheet sheet) {
        Assert.notNull(sheet, "sheet must be non-null");
        this.sheet = sheet;
    }

    private void sheetConfig(SheetConfig config) {
        Assert.notNull(config, "config must be non-null");
        this.listFirstRowNum = config.getListFirstRowNum();
        if (listFirstRowNum < 0) {
            findListFirstRowNum();
        }
        findColumnCountAndColumnNames();
        this.listLastRowNum = sheet.getLastRowNum();
        int configListLastRowNum = config.getListLastRowNum();
        if (configListLastRowNum < 0) {
            this.listLastRowNum += configListLastRowNum;
        } else if (configListLastRowNum < this.listLastRowNum) {
            this.listLastRowNum = configListLastRowNum;
        }
        this.isIgnoreException = config.isIgnoreException();
    }

    private void beanConfig(BeanConfig<?> config) {
        sheetConfig(config);
        this.beanConfig = config;
        this.fieldNames = config.getFieldNames();
        Assert.notEmpty(fieldNames, "fieldNames must be non-empty");
        this.targetClass = config.getTargetClass();
        Assert.notNull(targetClass, "targetClass must be non-null");
        //要处理的列数按最小的值
        if (fieldNames.length < columnCount) {
            this.columnCount = fieldNames.length;
        }
        this.fieldNameChains = new ArrayList<>(fieldNames.length);
        for (String fieldName : this.fieldNames) {
            this.fieldNameChains.add(Arrays.asList(fieldName.split("\\.")));
        }
        this.converterMap = config.getConverterMap();
    }

    private void parseMapList() throws CellException {
        rowNum = listFirstRowNum;
        for (; rowNum <= listLastRowNum; rowNum++) {
            row = sheet.getRow(rowNum);
            Map<Integer, Object> rowMap = new HashMap<>(columnCount, 1);
            for (columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                try {
                    cell = row.getCell(columnIndex);
                    rowMap.put(columnIndex, getCellValue());
                } catch (RuntimeException e) {
                    CellException cellException = new CellException(e, rowNum, columnIndex);
                    if (!isIgnoreException) {
                        throw cellException;
                    }
                    exceptions.add(cellException);
                }
            }
            mapList.add(rowMap);
        }
    }

    private void parseBeanList() throws InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException, NoSuchFieldException, CellException {
        parseBeanField();
        rowNum = listFirstRowNum;
        for (; rowNum <= listLastRowNum; rowNum++) {
            row = sheet.getRow(rowNum);
            Object baseObject = targetClass.getDeclaredConstructor().newInstance();
            for (columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                try {
                    List<Field> fieldChain = fieldChains.get(columnIndex);
                    if (fieldChain.isEmpty()) {
                        continue;
                    }
                    cell = row.getCell(columnIndex);
                    value = getCellValue();
                    stringValue = value == null ? null : value.toString();
                    ReflectionUtils.setFieldValue(baseObject, fieldChain, convert());
                } catch (RuntimeException e) {
                    CellException cellException = new CellException(e, rowNum, columnIndex);
                    if (!isIgnoreException) {
                        throw cellException;
                    }
                    exceptions.add(cellException);
                }
            }
            beanList.add(baseObject);
        }
    }

    public Object getCellValue() {
        switch (cell.getCellType()) {
            case NUMERIC:
                return cell.getNumericCellValue();
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case ERROR:
                return cell.getErrorCellValue();
            case _NONE:
            case BLANK:
            case STRING:
            case FORMULA:
            default:
                return cell.getStringCellValue();
        }
    }

    public Object convert() {
        if (converterMap.containsKey(fieldNames[columnIndex])) {
            return converterMap.get(fieldNames[columnIndex]).convert(this);
        } else {
            if (value == null) {
                return null;
            }
            List<Field> fieldChain = fieldChains.get(columnIndex);
            if (fieldChain.size() == 0) {
                return value;
            }
            Field targetField = fieldChain.get(fieldChain.size() - 1);
            if (targetField.getType().isAssignableFrom(value.getClass())) {
                return value;
            }
            Optional<CellConverter> optional = CellConverterRegistry.get(targetField.getType());
            if (optional.isPresent()) {
                return optional.get().convert(this);
            }
        }
        return value;
    }

    private void parseBeanField() throws NoSuchFieldException {
        fieldChains = new ArrayList<>(columnCount);
        for (int i = 0; i < columnCount; i++) {
            List<String> fieldNameChain = fieldNameChains.get(i);
            String firstFieldName = fieldNameChain.get(0);
            if (!"".equals(firstFieldName)) {
                fieldChains.add(Arrays.asList(ReflectionUtils.getFieldChain(targetClass, fieldNameChain)));
            } else {
                fieldChains.add(Collections.emptyList());
            }
        }
    }

    private void findListFirstRowNum() {
        listFirstRowNum = 0;
        int length = -1;
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            int l = sheet.getRow(i).getLastCellNum();
            if (l == length) {
                listFirstRowNum = i;
                return;
            }
            length = l;
        }
    }

    private void findColumnCountAndColumnNames() {
        columnCount = sheet.getRow(listFirstRowNum).getLastCellNum();
        if (listFirstRowNum > 0) {
            row = sheet.getRow(listFirstRowNum - 1);
            int l = row.getLastCellNum();
            columnNames = new String[l];
            for (int j = 0; j < l; j++) {
                cell = row.getCell(j);
                columnNames[j] = getCellValue().toString();
            }
        }
    }

    @Override
    public Object getValue() {
        return value;
    }

    @Override
    public String getStringValue() {
        return stringValue;
    }

    @Override
    public int getRowNum() {
        return rowNum;
    }

    @Override
    public int getColumnIndex() {
        return columnIndex;
    }

    @Override
    public List<Object> getBeanList() {
        return beanList;
    }

    @Override
    public XSSFRow getRow() {
        return row;
    }

    @Override
    public XSSFCell getCell() {
        return cell;
    }

    @Override
    public XSSFSheet getSheet() {
        return sheet;
    }

    @Override
    public BeanConfig<?> getBeanConfig() {
        return beanConfig;
    }
}
