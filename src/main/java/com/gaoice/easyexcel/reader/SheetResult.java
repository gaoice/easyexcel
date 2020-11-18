package com.gaoice.easyexcel.reader;

import com.gaoice.easyexcel.exception.CellException;

import java.util.List;

/**
 * 解析 Sheet 的结果
 *
 * @author gaoice
 */
public class SheetResult<T> {

    private List<T> result;

    /**
     * 解析的列名
     */
    private String[] columnNames;

    /**
     * 读取过程中发生的异常
     */
    private List<CellException> exceptions;

    public SheetResult() {
    }

    public SheetResult(List<T> result, String[] columnNames, List<CellException> exceptions) {
        this.result = result;
        this.columnNames = columnNames;
        this.exceptions = exceptions;
    }

    public String[] getColumnNames() {
        return columnNames;
    }

    public void setColumnNames(String[] columnNames) {
        this.columnNames = columnNames;
    }

    public List<T> getResult() {
        return result;
    }

    public void setResult(List<T> result) {
        this.result = result;
    }

    public List<CellException> getExceptions() {
        return exceptions;
    }

    public void setExceptions(List<CellException> exceptions) {
        this.exceptions = exceptions;
    }
}
