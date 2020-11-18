package com.gaoice.easyexcel.writer;

import com.gaoice.easyexcel.writer.sheet.SheetBuilder;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.util.List;

/**
 * @author gaoice
 */
public class ExcelWriter {

    public static void writeAndCloseOutputStream(SheetInfo sheetInfo, OutputStream out) throws Exception {
        writeOutputStream(sheetInfo, out);
        out.close();
    }

    public static void writeAndCloseOutputStream(List<SheetInfo> sheetInfos, OutputStream out) throws Exception {
        writeOutputStream(sheetInfos, out);
        out.close();
    }

    public static void writeOutputStream(SheetInfo sheetInfo, OutputStream out) throws Exception {
        SXSSFWorkbook workbook = createWorkbook(sheetInfo);
        workbook.write(out);
        workbook.dispose();
    }

    public static void writeOutputStream(List<SheetInfo> sheetInfos, OutputStream out) throws Exception {
        SXSSFWorkbook workbook = createWorkbook(sheetInfos);
        workbook.write(out);
        workbook.dispose();
    }

    public static SXSSFWorkbook createWorkbook(SheetInfo sheetInfo) throws Exception {
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        createSheet(workbook, sheetInfo);
        return workbook;
    }

    public static SXSSFWorkbook createWorkbook(List<SheetInfo> sheetInfos) throws Exception {
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        for (SheetInfo sheetInfo : sheetInfos) {
            createSheet(workbook, sheetInfo);
        }
        return workbook;
    }

    public static void createSheet(SXSSFWorkbook workbook, SheetInfo sheetInfo) throws IllegalAccessException, UnsupportedEncodingException, NoSuchFieldException {
        SheetBuilder.builder().buildSheet(workbook, sheetInfo);
    }

    public static SheetInfo builder() {
        return new SheetInfo();
    }
}
