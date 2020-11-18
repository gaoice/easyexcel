package com.gaoice.easyexcel.writer.style;

import com.gaoice.easyexcel.writer.sheet.SheetBuilderContext;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * 对 标题 列名 统计结果 的字体加粗
 *
 * @author gaoice
 */
public class BoldSheetStyle extends DefaultSheetStyle {

    @Override
    public CellStyle getTitleCellStyle(SheetBuilderContext<?> context) {
        if (titleCellStyle != null) {
            return titleCellStyle;
        }
        CellStyle style = super.getTitleCellStyle(context);
        context.getWorkbook().getFontAt(style.getFontIndexAsInt()).setBold(true);
        return style;
    }

    @Override
    public CellStyle getColumnNamesCellStyle(SheetBuilderContext<?> context) {
        if (columnNamesCellStyle != null) {
            return columnNamesCellStyle;
        }
        CellStyle style = super.getColumnNamesCellStyle(context);
        context.getWorkbook().getFontAt(style.getFontIndexAsInt()).setBold(true);
        return style;
    }

    @Override
    public CellStyle getColumnCountCellStyle(SheetBuilderContext<?> context) {
        if (columnCountCellStyle != null) {
            return columnCountCellStyle;
        }
        CellStyle style = super.getColumnCountCellStyle(context);
        context.getWorkbook().getFontAt(style.getFontIndexAsInt()).setBold(true);
        return style;
    }
}
