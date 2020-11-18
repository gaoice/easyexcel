package com.gaoice.easyexcel.reader.sheet;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.List;

/**
 * 单元格值包装类
 *
 * @author gaoice
 */
public interface SheetParserContext {

    Object getValue();

    String getStringValue();

    int getRowNum();

    int getColumnIndex();

    List<Object> getBeanList();

    XSSFRow getRow();

    XSSFCell getCell();

    XSSFSheet getSheet();

    BeanConfig<?> getBeanConfig();
}
