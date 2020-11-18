package com.gaoice.easyexcel.reader;

import com.gaoice.easyexcel.reader.sheet.BeanConfig;
import com.gaoice.easyexcel.reader.sheet.SheetConfig;
import com.gaoice.easyexcel.reader.sheet.SheetParser;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.List;
import java.util.Map;

/**
 * @author gaoice
 */
public class ExcelReader {

    public static SheetResult<Map<Integer, Object>> parseSheetResult(String fileName) throws IOException {
        try (FileInputStream inputStream = new FileInputStream(fileName)) {
            return parseSheetResult(inputStream);
        }
    }

    public static <T> SheetResult<T> parseSheetResult(String fileName, BeanConfig<T> config) throws IOException, NoSuchMethodException, NoSuchFieldException, InstantiationException, IllegalAccessException, InvocationTargetException {
        try (FileInputStream inputStream = new FileInputStream(fileName)) {
            return parseSheetResult(inputStream, config);
        }
    }

    public static SheetResult<Map<Integer, Object>> parseSheetResult(InputStream inputStream) throws IOException {
        return parseSheetResult(inputStream, new SheetConfig());
    }

    public static SheetResult<Map<Integer, Object>> parseSheetResult(InputStream inputStream, SheetConfig config) throws IOException {
        return parseSheetResult(new XSSFWorkbook(inputStream), config);
    }

    public static <T> SheetResult<T> parseSheetResult(InputStream inputStream, BeanConfig<T> config) throws IOException, InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException, NoSuchFieldException {
        return parseSheetResult(new XSSFWorkbook(inputStream), config);
    }

    public static SheetResult<Map<Integer, Object>> parseSheetResult(XSSFWorkbook workbook, SheetConfig config) {
        return parseSheetResult(workbook.getSheetAt(0), config);
    }

    public static <T> SheetResult<T> parseSheetResult(XSSFWorkbook workbook, BeanConfig<T> config) throws NoSuchMethodException, NoSuchFieldException, InstantiationException, IllegalAccessException, InvocationTargetException {
        return parseSheetResult(workbook.getSheetAt(0), config);
    }

    public static SheetResult<Map<Integer, Object>> parseSheetResult(XSSFSheet sheet, SheetConfig config) {
        return SheetParser.parser().parseMapList(sheet, config);
    }

    public static <T> SheetResult<T> parseSheetResult(XSSFSheet sheet, BeanConfig<T> config) throws InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException, NoSuchFieldException {
        return SheetParser.parser().parseBeanList(sheet, config);
    }

    // parseList

    public static List<Map<Integer, Object>> parseList(String fileName) throws IOException {
        try (FileInputStream inputStream = new FileInputStream(fileName)) {
            return parseList(inputStream);
        }
    }

    public static <T> List<T> parseList(String fileName, BeanConfig<T> config) throws IOException, NoSuchMethodException, NoSuchFieldException, InstantiationException, IllegalAccessException, InvocationTargetException {
        try (FileInputStream inputStream = new FileInputStream(fileName)) {
            return parseList(inputStream, config);
        }
    }

    public static List<Map<Integer, Object>> parseList(InputStream inputStream) throws IOException {
        return parseList(inputStream, new SheetConfig());
    }

    public static List<Map<Integer, Object>> parseList(InputStream inputStream, SheetConfig config) throws IOException {
        return parseList(new XSSFWorkbook(inputStream), config);
    }

    public static <T> List<T> parseList(InputStream inputStream, BeanConfig<T> config) throws IOException, InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException, NoSuchFieldException {
        return parseList(new XSSFWorkbook(inputStream), config);
    }

    public static List<Map<Integer, Object>> parseList(XSSFWorkbook workbook, SheetConfig config) {
        return parseList(workbook.getSheetAt(0), config);
    }

    public static <T> List<T> parseList(XSSFWorkbook workbook, BeanConfig<T> config) throws NoSuchMethodException, NoSuchFieldException, InstantiationException, IllegalAccessException, InvocationTargetException {
        return parseList(workbook.getSheetAt(0), config);
    }

    public static List<Map<Integer, Object>> parseList(XSSFSheet sheet, SheetConfig config) {
        return SheetParser.parser().parseMapList(sheet, config).getResult();
    }

    public static <T> List<T> parseList(XSSFSheet sheet, BeanConfig<T> config) throws InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException, NoSuchFieldException {
        return SheetParser.parser().parseBeanList(sheet, config).getResult();
    }
}
