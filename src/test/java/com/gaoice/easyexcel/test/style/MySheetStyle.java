package com.gaoice.easyexcel.test.style;

import com.gaoice.easyexcel.SheetInfo;
import com.gaoice.easyexcel.style.BoldSheetStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;


public class MySheetStyle extends BoldSheetStyle {
    private CellStyle warnStyle;
    private SheetInfo sheetInfo;

    /**
     * 自定义构造方法，传入所需变量
     */
    public MySheetStyle(SheetInfo sheetInfo) {
        this.sheetInfo = sheetInfo;
    }

    /**
     * 对 分数列小于60分的单元格 设置红色字体
     */
    @Override
    public CellStyle getListCellStyle(SXSSFWorkbook workbook, int listIndex, int columnIndex, Object v) {
        String[] names = sheetInfo.getFieldNames();
        if ((names[columnIndex].equals("grade.chineseGrade")
                || names[columnIndex].equals("grade.mathGrade")
                || names[columnIndex].equals("grade.englishGrade"))
                && (v == null || Double.parseDouble(v.toString()) < 60)) {
            if (warnStyle == null) {
                warnStyle = newSimpleStyle(workbook);
                Font font = workbook.getFontAt(warnStyle.getFontIndexAsInt());
                //(short) 10代表红色
                font.setColor((short) 10);
                warnStyle.setWrapText(false);
            }
            return warnStyle;
        }
        return super.getListCellStyle(workbook, listIndex, columnIndex, v);
    }

    /**
     * 长度处理，身份证列自适应长度会有少许遮挡，在这里处理补足
     */
    @Override
    public void handleColumnMaxBytesLength(int[] columnMaxBytesLength) {
        String[] names = sheetInfo.getFieldNames();
        for (int i = 0; i < names.length; i++) {
            if (names[i].equals("cardId")) {
                columnMaxBytesLength[i] += 4;
            }
        }
        /*
         * 父类方法对超过255的进行限制，否则会报错
         */
        super.handleColumnMaxBytesLength(columnMaxBytesLength);
    }

    /**
     * 清除样式
     */
    @Override
    public void clearStyleCache() {
        super.clearStyleCache();
        warnStyle = null;
    }
}
