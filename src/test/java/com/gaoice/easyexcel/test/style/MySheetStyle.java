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
     * 对分数列小于60分的单元格设置红色字体
     *
     * @param workbook
     * @param listIndex
     * @param columnIndex
     * @param v
     * @return
     */
    @Override
    public CellStyle getListCellStyle(SXSSFWorkbook workbook, int listIndex, int columnIndex, Object v) {
        String[] names = sheetInfo.getClassFieldNames();
        if ((names[columnIndex].equals("grade.chineseGrade")
                || names[columnIndex].equals("grade.mathGrade")
                || names[columnIndex].equals("grade.englishGrade"))
                && (Double.valueOf(v.toString()) < 60)) {
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
     *
     * @param columnMaxBytesLength
     */
    @Override
    public void columnMaxBytesLengthHandler(int[] columnMaxBytesLength) {
        String[] names = sheetInfo.getClassFieldNames();
        for (int i = 0; i < names.length; i++) {
            if (names[i].equals("cardId")) {
                columnMaxBytesLength[i] += 4;
            }
        }
        /*
         * 父类方法对超过255的进行限制，否则会报错
         */
        super.columnMaxBytesLengthHandler(columnMaxBytesLength);
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
