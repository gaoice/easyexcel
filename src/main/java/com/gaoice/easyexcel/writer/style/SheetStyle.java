package com.gaoice.easyexcel.writer.style;

import com.gaoice.easyexcel.writer.sheet.SheetBuilderContext;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * 一个 workbook 能够创建的 CellStyle 数量有限制
 * 所以对于每个返回 CellStyle 的方法，在第二次调用该方法的时候应该尽量使用缓存的 CellStyle
 *
 * @author gaoice
 */
public interface SheetStyle {

    /**
     * 获取 标题 样式
     *
     * @param context 当前上下文
     *                使用 {@link SheetBuilderContext#getWorkbook()} 获取 workbook
     *                CellStyle 需要调用 workbook.createCellStyle() 创建出来
     * @return cell style
     */
    CellStyle getTitleCellStyle(SheetBuilderContext<?> context);

    /**
     * 获取 列名（表头） 样式，每列的列名字都会调用此方法设置样式
     *
     * @param context 使用 {@link SheetBuilderContext#getColumnIndex()} 获取列索引序号
     *                可以通过序号判断，对特殊的列使用特殊的样式
     * @return cell style
     */
    CellStyle getColumnNamesCellStyle(SheetBuilderContext<?> context);

    /**
     * 获取 表格 样式，每个 list 中元素字段对应的单元格都会调用此方法设置样式
     *
     * @param context 使用 {@link SheetBuilderContext#getListIndex()} 获取索引
     *                可以通过索引的奇偶，设置颜色相间的表格样式
     *                由于 list 对应的单元格可能会较多，这个方法会调用多次，相同的样式应该缓存起来重复使用
     * @return cell style
     */
    CellStyle getListCellStyle(SheetBuilderContext<?> context);

    /**
     * 获取 统计结果行 样式，每个统计结果都会调用方法设置样式
     *
     * @param context 使用 context.getWorkbook() 获取 workbook
     *                CellStyle 需要调用 workbook.createCellStyle() 创建出来
     *                由于 list 对应的单元格可能会较多，这个方法会调用多次，相同的样式应该缓存起来重复使用
     * @return cell style
     */
    CellStyle getColumnCountCellStyle(SheetBuilderContext<?> context);

    /**
     * 列最大字节长度处理，用于单元格宽度会自适应，只会调用一次，需在一次处理中处理所有列宽
     *
     * @param context 使用 {@link SheetBuilderContext#getColumnMaxByteLengths()} 获取列宽数组进行处理
     */
    void handleColumnMaxBytesLength(SheetBuilderContext<?> context);

    /**
     * 不同 workbook style 不能通用
     * SheetBuilder 在工作完成后将会调用此方法清除 style 缓存
     * 没有重复使用 SheetStyle 不必实现此方法
     */
    default void clearStyleCache() {
    }
}
