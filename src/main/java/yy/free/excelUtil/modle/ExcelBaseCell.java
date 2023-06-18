package yy.free.excelUtil.modle;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import yy.free.excelUtil.ExcelUtils;


import java.util.HashMap;
import java.util.Map;

/**
 * 构建cell 可操作对象，
 * 避免在外部pdf 中操作cell 避免类名冲突
 *
 * @author denny.gong
 * @version 1.0
 * @date 2022/1/25 10:02
 **/
@ToString
@Setter
@Getter
public class ExcelBaseCell {

    /**
     * 是否是合并单元格
     */
    private boolean isMergedRegion;

    /**
     * 单元格类型
     */
    private CellType cellType;

    /**
     * 单元格值
     */
    private Object value;

    /**
     * 行号
     */
    private int row;

    /**
     * 列号
     */
    private int column;

    /**
     * 区块横跨列
     */
    private int columnSpan = 1;

    /**
     * 区块横跨行
     */
    private int rowSpan = 1;

    /**
     * map对象 用于存放样式
     * alignment 设置单元格水平方向对其方式
     * bold 是否加粗
     * font_red 字体三原色：红
     * font_green 字体三原色：绿
     * font_blue 字体三原色：蓝
     * topBorderColor 顶边框颜色
     * bottomColor 底边框颜色
     * leftColor 左边框颜色
     * rightColor 右边框颜色
     * backgroundColor 背景色（Excel填充色为前景色，pdf只有背景色）
     * borderTop 顶边框宽度
     * borderBottom 底边框宽度
     * borderLeft 左边框宽度
     * borderRight 右边框宽度
     */
    private Map<String, String> stlys;

    /**
     * 初始化
     *
     * @param isMergedRegion 是否合并单元格
     * @param row            行
     * @param column         列
     */
    public ExcelBaseCell(boolean isMergedRegion, int row, int column) {
        this.isMergedRegion = isMergedRegion;
        this.cellType = CellType.BLANK;
        this.row = row;
        this.column = column;
    }

    public ExcelBaseCell(boolean isMergedRegion, int row, int column, Cell cell, CellRangeAddress address) {
        this.isMergedRegion = isMergedRegion;
        this.row = row;
        this.column = column;
        if (ExcelUtils.isEmptyCell(cell)) {
            // 获取单元格样式
            stlys = getStyles(cell);
            this.cellType = CellType.BLANK;
        } else {
            this.cellType = cell.getCellType();
            switch (cell.getCellType()) {
                case NUMERIC:
                    this.value = ExcelUtils.getNumericCellToString(cell);
                    break;
                case STRING:
                    this.value = cell.getStringCellValue();
                    break;
                case FORMULA:
                    String valueStr;
                    try {
                        valueStr = String.format("%.2f", cell.getNumericCellValue());
                    } catch (IllegalStateException e) {
                        valueStr = "0";
                    }
                    this.value = valueStr;
                    break;
                case BLANK:
                    break;
                case BOOLEAN:
                    this.value = cell.getBooleanCellValue();
                default:
                    break;
            }
            // 获取单元格样式
            stlys = getStyles(cell);
        }
        if (address != null) {
            this.rowSpan = Math.max(address.getLastRow() - address.getFirstRow(), 1);
            this.columnSpan = Math.max(address.getLastColumn() - address.getFirstColumn(), 1);

            if (address.getLastRow() != address.getFirstRow()) {
                this.rowSpan = Math.max(address.getLastRow() - address.getFirstRow(), 1) + 1;
            }
            if (address.getLastColumn() != address.getFirstColumn()) {
                this.columnSpan = Math.max(address.getLastColumn() - address.getFirstColumn(), 1) + 1;
            }
        }
    }

    /**
     * 获取单元格样式
     *
     * @param cell 单元格类型
     * @return map样式
     */
    private Map<String, String> getStyles(Cell cell) {
        Map<String, String> map = new HashMap<>();
        CellStyle style = cell.getCellStyle();
        Workbook workbook = cell.getSheet().getWorkbook();
        // 设置单元格水平方向对其方式
        map.put("alignment", String.valueOf(style.getAlignment().getCode()));
        // 获取字体三原色
        ExcelUtils.setFontStyle(workbook, style.getFontIndexAsInt(), map);
        // 获取边框样式
        ExcelUtils.setBorderColor(cell, map);
        return map;
    }
}
