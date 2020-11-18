package com.gaoice.easyexcel.exception;

/**
 * 处理单元格值发生的异常
 *
 * @author gaoice
 */
public class CellException extends RuntimeException {

    public CellException() {
        super();
    }

    public CellException(String message) {
        super(message);
    }

    public CellException(Throwable cause) {
        super(cause);
    }

    public CellException(String message, Throwable cause) {
        super(message, cause);
    }

    public CellException(Throwable cause, int i, int j) {
        super("[rowIndex:" + i + ",columnIndex:" + j + "]" + cause.getMessage(), cause);
    }
}
