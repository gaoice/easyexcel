package com.gaoice.easyexcel.reader.sheet;

/**
 * 要读取的 sheet 的配置
 *
 * @author gaoice
 */
public class SheetConfig {

    /**
     * List 开始的行号，从 0 开始计数
     * 值小于 0 为自动判断，自动判断会忽略标题和表头
     */
    private int listFirstRowNum = -1;

    /**
     * List 结束的行号
     * 正数举例：
     * 1 读取到第 1 行（从 0 开始计数），包含第一行，以此类推
     * 负数举例：
     * -1 读取至倒数第一行，不包含倒数第一行
     * -2 读取至倒数第二行，不包含倒数第一二行，以此类推
     * <p>
     * 要读完整个文档，使用 Integer.MAX_VALUE 即可
     */
    private int listLastRowNum = Integer.MAX_VALUE;

    /**
     * 是否忽略处理单元格值发生的异常
     */
    private boolean isIgnoreException = false;

    public int getListFirstRowNum() {
        return listFirstRowNum;
    }

    public SheetConfig setListFirstRowNum(int listFirstRowNum) {
        this.listFirstRowNum = listFirstRowNum;
        return this;
    }

    public int getListLastRowNum() {
        return listLastRowNum;
    }

    public SheetConfig setListLastRowNum(int listLastRowNum) {
        this.listLastRowNum = listLastRowNum;
        return this;
    }

    public boolean isIgnoreException() {
        return isIgnoreException;
    }

    public SheetConfig setIgnoreException(boolean ignoreException) {
        isIgnoreException = ignoreException;
        return this;
    }
}
