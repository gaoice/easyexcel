package com.gaoice.easyexcel.test.writer.style;

import com.gaoice.easyexcel.writer.sheet.SheetBuilderContext;
import com.gaoice.easyexcel.writer.style.BoldSheetStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;


public class MySheetStyle extends BoldSheetStyle {
    private CellStyle warnStyle;

    /**
     * 对 分数列小于60分的单元格 设置红色字体
     */
    @Override
    public CellStyle getListCellStyle(SheetBuilderContext<?> context) {
        int columnIndex = context.getColumnIndex();
        Object value = context.getValue();
        String[] names = context.getSheetInfo().getFieldNames();
        if ((names[columnIndex].equals("grade.chineseGrade")
                || names[columnIndex].equals("grade.mathGrade")
                || names[columnIndex].equals("grade.englishGrade"))
                && (value == null || Double.parseDouble(value.toString()) < 60)) {
            if (warnStyle == null) {
                warnStyle = newSimpleStyle(context);
                Font font = context.getWorkbook().getFontAt(warnStyle.getFontIndexAsInt());
                //(short) 10代表红色
                font.setColor((short) 10);
                warnStyle.setWrapText(false);
            }
            return warnStyle;
        }
        return super.getListCellStyle(context);
    }

    /**
     * 长度处理，身份证列自适应长度会有少许遮挡，在这里处理补足
     */
    @Override
    public void handleColumnMaxBytesLength(SheetBuilderContext<?> context) {
        String[] fieldNames = context.getSheetInfo().getFieldNames();
        int[] columnMaxByteLengths = context.getColumnMaxByteLengths();
        for (int i = 0; i < fieldNames.length; i++) {
            if (fieldNames[i].equals("cardId")) {
                columnMaxByteLengths[i] += 5;
            }
        }
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
