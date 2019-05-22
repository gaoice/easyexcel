package com.gaoice.easyexcel;

import com.gaoice.easyexcel.data.Converter;
import com.gaoice.easyexcel.data.Counter;
import com.gaoice.easyexcel.style.DefaultSheetStyle;
import com.gaoice.easyexcel.style.SheetStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.ByteArrayOutputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.*;

public class ExcelBuilder {

    private static char virtualFieldStart = '#';
    private static String charset = "utf-8";

    public static XSSFWorkbook createWorkbook(SheetInfo sheetInfo) throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
        createSheet(workbook, sheetInfo);
        return workbook;
    }

    public static XSSFWorkbook createWorkbook(List<SheetInfo> sheetInfos) throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
        for (SheetInfo sheetInfo : sheetInfos) {
            createSheet(workbook, sheetInfo);
        }
        return workbook;
    }

    public static OutputStream createOutputStream(SheetInfo sheetInfo) throws Exception {
        XSSFWorkbook workbook = createWorkbook(sheetInfo);
        OutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        return out;
    }

    public static OutputStream createOutputStream(List<SheetInfo> sheetInfos) throws Exception {
        XSSFWorkbook workbook = createWorkbook(sheetInfos);
        OutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        return out;
    }

    public static XSSFSheet createSheet(XSSFWorkbook workbook, SheetInfo sheetInfo) throws Exception {
        String sheetName = sheetInfo.getSheetName();
        String title = sheetInfo.getTitle();
        String[] columnNames = sheetInfo.getColumnNames();
        String[] classFieldNames = sheetInfo.getClassFieldNames();
        List<Object> list = sheetInfo.getList();
        SheetStyle sheetStyle = sheetInfo.getSheetStyle();
        Map<String, Object> converterMap = sheetInfo.getConverterMap();
        Map<String, Object> counterMap = sheetInfo.getCounterMap();

        if (classFieldNames == null) {
            throw new IllegalArgumentException(sheetName + ": classFieldNames shall not be null");
        }
        if (sheetStyle == null) {
            sheetStyle = new DefaultSheetStyle();
        }
        List<List<String>> classFieldNameBunchs = new ArrayList<>();
        for (String classFieldName : classFieldNames) {
            classFieldNameBunchs.add(Arrays.asList(classFieldName.split("\\.")));
        }
        List<List<Field>> classFieldBunchs;

        final int columnCount = classFieldNames.length;
        int[] columnMaxLength = new int[columnCount];
        Object[] countResults = counterMap == null ? null : new Object[columnCount];

        int nextRowNum = 0;

        XSSFSheet sheet = workbook.createSheet(sheetName);
        XSSFRow row;
        XSSFCell cell;
        /*
         * title
         */
        if (title != null) {
            sheet.addMergedRegion(new CellRangeAddress(nextRowNum, nextRowNum, 0, columnCount - 1));
            row = sheet.createRow(nextRowNum++);
            cell = row.createCell(0);
            cell.setCellValue(title);
            cell.setCellStyle(sheetStyle.getTitleCellStyle(workbook));
        }
        /*
         * columnNames
         */
        if (columnNames != null) {
            if (columnNames.length != classFieldNames.length) {
                throw new IllegalArgumentException(sheetName + ": columnNames.length shall equals classFieldNames.length");
            }
            row = sheet.createRow(nextRowNum++);
            for (int i = 0; i < columnCount; i++) {
                cell = row.createCell(i);
                cell.setCellValue(columnNames[i]);
                cell.setCellStyle(sheetStyle.getColumnNamesCellStyle(workbook, i));
                int valueLength = columnNames[i].getBytes(charset).length;
                if (valueLength > columnMaxLength[i]) {
                    columnMaxLength[i] = valueLength;
                }
            }
        }

        if (list != null && list.size() > 0) {
            /*
             * field
             */
            if (sheetInfo.getFieldCache() != null) {
                classFieldBunchs = sheetInfo.getFieldCache();
            } else {
                classFieldBunchs = new ArrayList<>();
                Class baseClass = list.get(0).getClass();
                for (int i = 0; i < columnCount; i++) {
                    Class c = baseClass;
                    List<String> classFieldNameBunch = classFieldNameBunchs.get(i);
                    List<Field> classFieldBunch = new ArrayList<>();
                    String firstClassFieldName = classFieldNameBunch.get(0);
                    if (!firstClassFieldName.equals("") && firstClassFieldName.charAt(0) != virtualFieldStart) {
                        for (int j = 0; j < classFieldNameBunch.size(); j++) {
                            Field f = c.getDeclaredField(classFieldNameBunch.get(j));
                            f.setAccessible(true);
                            classFieldBunch.add(f);
                            if (j != (classFieldNameBunch.size() - 1)) {
                                c = Class.forName(f.getGenericType().getTypeName());
                            }
                        }
                    }
                    classFieldBunchs.add(classFieldBunch);
                }
                sheetInfo.setFieldCache(classFieldBunchs);
            }
            /*
             * list fill
             */
            for (int i = 0; i < list.size(); i++) {
                row = sheet.createRow(nextRowNum++);
                Object baseObject = list.get(i);
                for (int j = 0; j < columnCount; j++) {
                    Object o = baseObject;
                    for (Field f : classFieldBunchs.get(j)) {
                        if (o != null) {
                            o = f.get(o);
                        } else {
                            break;
                        }
                    }
                    Object originalValue = o;
                    /*
                     * convert
                     */
                    String classFieldName = classFieldNames[j];
                    if (converterMap != null && converterMap.containsKey(classFieldName)) {
                        Object converter = converterMap.get(classFieldName);
                        if (converter instanceof Converter) {
                            Converter cellConvert = (Converter) converter;
                            o = cellConvert.convert(sheetInfo, o, i, j);
                        } else if (converter instanceof Map) {
                            Map dictionary = (Map) converter;
                            o = dictionary.get(o);
                        }
                    }
                    /*
                     * count
                     */
                    if (counterMap != null && counterMap.containsKey(classFieldName)) {
                        Object counter = counterMap.get(classFieldName);
                        if (counter instanceof Counter) {
                            Counter columnCounter = (Counter) counter;
                            countResults[j] = columnCounter.count(sheetInfo, originalValue, i, j, countResults[j]);
                        } else if (o instanceof BigDecimal) {
                            BigDecimal o1 = (BigDecimal) o;
                            BigDecimal r = (BigDecimal) countResults[j];
                            countResults[j] = r == null ? o1 : o1.add(r);
                        } else if (o instanceof Number) {
                            Number o1 = (Number) o;
                            Number r = (Number) countResults[j];
                            countResults[j] = r == null ? o1 : (o1.doubleValue() + r.doubleValue());
                        } else {
                            countResults[j] = "not support default counter";
                        }
                    }
                    /*
                     * set cell value and style
                     */
                    String value = o == null ? "" : o.toString();
                    cell = row.createCell(j);
                    if (o instanceof Number) {
                        Number num = (Number) o;
                        cell.setCellValue(num.doubleValue());
                    } else {
                        cell.setCellValue(value);
                    }
                    cell.setCellStyle(sheetStyle.getListCellStyle(workbook, i, j, o));
                    int valueLength = value.getBytes(charset).length;
                    if (valueLength > columnMaxLength[j]) {
                        columnMaxLength[j] = valueLength;
                    }
                }
            }
        }

        /*
         * create count result row
         */
        if (countResults != null) {
            row = sheet.createRow(nextRowNum++);
            for (int i = 0; i < columnCount; i++) {
                String value = countResults[i] == null ? "" : countResults[i].toString();
                cell = row.createCell(i);
                cell.setCellValue(value);
                cell.setCellStyle(sheetStyle.getColumnCountCellStyle(workbook, i, countResults[i]));
                int valueLength = value.getBytes(charset).length;
                if (valueLength > columnMaxLength[i]) {
                    columnMaxLength[i] = valueLength;
                }
            }
        }

        /*
         * set column width
         */
        sheetStyle.columnMaxBytesLengthHandler(columnMaxLength);
        for (int i = 0; i < columnCount; i++) {
            sheet.setColumnWidth(i, columnMaxLength[i] * 256);
        }
        /*
         * clear sheet style cache
         */
        sheetStyle.clearStyleCache();
        return sheet;
    }

}
