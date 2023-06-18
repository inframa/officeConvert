package yy.free.excelUtil;

import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.format.CellFormat;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import yy.free.excelUtil.modle.ExcelBaseCell;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * Excel工具类
 * <p>
 *
 * </p>
 *
 * @author yy
 */
public class ExcelUtils {

    /**
     * 获取最大列值
     *
     * @param sheet 子表
     */
    public static int getSheetCellNumn(Sheet sheet) {
        int maxCells = 0;
        for (int row = 0; row < sheet.getLastRowNum(); row++) {
            if (!isEmptyRow(sheet.getRow(row))) {
                maxCells = Math.max(sheet.getRow(row).getLastCellNum(), maxCells);
            }
        }
        return maxCells;
    }

    /**
     * 判定行是否为空行
     *
     * @param row 行信息
     * @return 是否为空
     */
    public static boolean isEmptyRow(Row row) {
        return null == row || row.toString().isEmpty();
    }

    /**
     * 判定单元格信息是否为空
     *
     * @param cell 单元格信息
     * @return 是否为空
     */
    public static boolean isEmptyCell(Cell cell) {
        return null == cell || CellType.BLANK == cell.getCellType();
    }


    /**
     * 转换excel 日期数值类型
     *
     * @param cell 单元格信息
     * @return 类型
     */

    public static Object getNumericCellToString(Cell cell) {
        String cellValue = "";
        //处理yyyy年m月d日,h时mm分,yyyy年m月,m月d日等含文字的日期格式
        //判断cell.getCellStyle().getDataFormat()值，解析数值格式
                /*
                    yyyy-MM-dd----- 14
                    HH:mm:ss ---------21
                    yyyy-MM-dd HH:mm:ss ---------22
                    yyyy年m月d日--- 31
                    h时mm分 ------- 32
                    yyyy年m月------- 57
                    m月d日 ---------- 58
                    HH:mm----------- 20
                */
        // 获取日期格式
        short format = cell.getCellStyle().getDataFormat();
        Date date;
        double value;
        SimpleDateFormat sdf;
        switch (format) {
            case 0:
                //处理数值格式
                if ("General".equals(cell.getCellStyle().getDataFormatString())) {
                    DataFormatter dataFormatter = new DataFormatter();
                    cellValue = dataFormatter.formatCellValue(cell);
                } else {
                    cellValue = CellFormat.getInstance(cell.getCellStyle().getDataFormatString()).apply(cell).text;
                }
                break;
            case 14:
                date = cell.getDateCellValue();
                cellValue = DateFormatUtils.format(date, "yyyy-MM-dd");
                break;
            case 20:
                sdf = new SimpleDateFormat("HH:mm");
                value = cell.getNumericCellValue();
                date = DateUtil.getJavaDate(value);
                cellValue = sdf.format(date);
                break;
            case 21:
                date = cell.getDateCellValue();
                cellValue = DateFormatUtils.format(date, "HH:mm:ss");
                break;
            case 22:
                date = cell.getDateCellValue();
                cellValue = DateFormatUtils.format(date, "yyyy-MM-dd HH:mm:ss");
                break;
            case 31:
                date = cell.getDateCellValue();
                cellValue = DateFormatUtils.format(date, "yyyy年M月d日");
                break;
            case 32:
                sdf = new SimpleDateFormat("H时mm分");
                value = cell.getNumericCellValue();
                date = DateUtil.getJavaDate(value);
                cellValue = sdf.format(date);
                break;
            case 57:
                date = cell.getDateCellValue();
                cellValue = DateFormatUtils.format(date, "yyyy年M月");
                break;
            case 58:
                date = cell.getDateCellValue();
                cellValue = DateFormatUtils.format(date, "M月d日");
                break;
            case 177:
                cellValue = String.format("%.2f", cell.getNumericCellValue());
                break;
            default:
                cellValue = CellFormat.getInstance(cell.getCellStyle().getDataFormatString()).apply(cell).text;
                break;
        }
        if (cell.toString().contains("%")) {
            // 判断是否是百分数类型
            cellValue = cell.getNumericCellValue() * 100 + "";
        }
        return cellValue;
    }

    /**
     * 构造可操作pdfCell对象
     *
     * @param list   合并对象值
     * @param row    目标单元个行
     * @param column 目标单元个列
     * @return 构建基础 pdfCell
     */
    public static ExcelBaseCell getMergedRegion(List<CellRangeAddress> list, int row, int column, Cell cell) {
        for (CellRangeAddress range : list) {
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            //如果目标行编号大于等于合并单元格的起始行
            //并且目标行编号小于等于
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    if (row == firstRow && column == firstColumn) {
                        ExcelBaseCell pdfCell = new ExcelBaseCell(true, row, column, cell, range);
                        return pdfCell;
                    }
                    return null;
                }
            }
        }
        return new ExcelBaseCell(false, row, column, cell, null);
    }

    /**
     * 构造可操作pdfCell对象集合
     *
     * @param sheet    表单
     * @param maxColum 最大列
     * @return 可操作pdfCell对象集合
     */
    public static List<ExcelBaseCell> imptBasePdfCell(Sheet sheet, int maxColum) {
        List<ExcelBaseCell> list = new ArrayList<>();
        //获取合并单元格集合
        List<CellRangeAddress> cellRanges = getRange(sheet);
        //遍历行并处理cell 值
        //getLastRowNum更换为getPhysicalNumberOfRows
        for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
            Row row = sheet.getRow(i);
            //如果该行为空则构建相关maxcolum 插入
            if (isEmptyRow(row)) {
                list.addAll(setEmptyList(i, maxColum));
                continue;
            }
            //遍历循环处理行中包含的列值
            for (int j = 0; j < maxColum; ++j) {
                Cell cell = row.getCell(j);
                ExcelBaseCell pdfCell = getMergedRegion(cellRanges, i, j, cell);
                if (pdfCell != null) {
                    list.add(pdfCell);
                }
            }

        }
        return list;
    }

    /**
     * 构建空行
     *
     * @param row        行号
     * @param maxColumns 跨越的列值
     * @return List<ExcelBaseCell>
     */
    public static List<ExcelBaseCell> setEmptyList(int row, int maxColumns) {
        List<ExcelBaseCell> basePdfCells = new ArrayList<>();
        for (int i = 0; i < maxColumns; i++) {
            basePdfCells.add(new ExcelBaseCell(false, row, i, null, null));
        }
        return basePdfCells;
    }

    /**
     * 获取合并单元格集合
     *
     * @param sheet 表单
     * @return 合并单元格集合
     */
    public static List<CellRangeAddress> getRange(Sheet sheet) {
        List<CellRangeAddress> list = new ArrayList<>();
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            list.add(range);
        }
        return list;
    }

    /**
     * 只有当单元格只有一种字体时，这才有效。如果单元格包含富文本字符串，则每个格式化文本串都有相应的字体。然后需要获取并遍历RichTextString。这要复杂得多，需要为HSSF和XSSF做不同的事情
     *
     * @param workbook 当前工作界面
     * @param fontIdx  字顺序
     */
    public static void setFontStyle(Workbook workbook, int fontIdx, Map<String, String> map) {
        Font font = workbook.getFontAt(fontIdx);
        // RGB = [0,0,0] 为黑色
        int red = 0;
        int green = 0;
        int blue = 0;
        if (font instanceof HSSFFont) {
            HSSFColor color = ((HSSFFont) font).getHSSFColor((HSSFWorkbook) workbook);
            // 为空默认为黑色
            if (color != null) {
                // rgb[0]: red, rgb[1]: green, rgb[2]: blue
                short[] rgb = color.getTriplet();
                red = rgb[0];
                green = rgb[1];
                blue = rgb[2];
            }
        } else if (font instanceof XSSFFont) {
            XSSFColor xc = ((XSSFFont) font).getXSSFColor();
            byte[] data = new byte[]{0, 1, 2};
            if (null == xc) {
                data[0] = 0 & 0xFF;
                data[1] = 0 & 0xFF;
                data[2] = 0 & 0xFF;
            } else {
                if (xc.getTint() != 0.0) {
                    data = xc.getRGBWithTint();
                } else {
                    data = xc.getARGB();
                }
                if (data == null) {
                    data[0] = 0 & 0xFF;
                    data[1] = 0 & 0xFF;
                    data[2] = 0 & 0xFF;
                }
            }
            int idx = 0;
            int alpha = 255;
            if (data.length == 4) {
                alpha = data[idx++] & 0xFF;
            }
            red = data[idx++] & 0xFF;
            green = data[idx++] & 0xFF;
            blue = data[idx++] & 0xFF;
        }
        map.put("font_red", String.valueOf(red));
        map.put("font_green", String.valueOf(green));
        map.put("font_blue", String.valueOf(blue));
        if (font.getBold()) {
            map.put("bold", "1");
        } else {
            map.put("bold", "0");
        }

    }


    /**
     * 设置单元格边框颜色
     *
     * @param cell 单元格对象
     * @param map  map
     */
    public static void setBorderColor(Cell cell, Map<String, String> map) {
        CellStyle style = cell.getCellStyle();
        Workbook workbook = cell.getSheet().getWorkbook();
        if (style instanceof XSSFCellStyle) {
            // 顶边框颜色
            XSSFColor topColor = ((XSSFCellStyle) style).getTopBorderXSSFColor();
            map.put("topBorderColor", getXssfRgb(topColor, false));
            // 底边框颜色
            XSSFColor bottomColor = ((XSSFCellStyle) style).getBottomBorderXSSFColor();
            map.put("bottomColor", getXssfRgb(bottomColor, false));
            // 左边框颜色
            XSSFColor leftColor = ((XSSFCellStyle) style).getLeftBorderXSSFColor();
            map.put("leftColor", getXssfRgb(leftColor, false));
            // 右边框颜色
            XSSFColor rightColor = ((XSSFCellStyle) style).getRightBorderXSSFColor();
            map.put("rightColor", getXssfRgb(rightColor, false));
            // WPS单元格填充色的前景色，office没测试
            XSSFColor backgroundColor = ((XSSFCellStyle) style).getFillForegroundXSSFColor();
            map.put("backgroundColor", getXssfRgb(backgroundColor, true));
        } else if (style instanceof HSSFCellStyle) {
            // 顶边框颜色
            HSSFColor topColor = ((HSSFWorkbook) workbook).getCustomPalette().getColor(((HSSFCellStyle) style).getTopBorderColor());
            map.put("topBorderColor", getHssfRgb(topColor, false));
            // 底边框颜色
            HSSFColor bottomColor = ((HSSFWorkbook) workbook).getCustomPalette().getColor(((HSSFCellStyle) style).getBottomBorderColor());
            map.put("bottomColor", getHssfRgb(bottomColor, false));
            // 左边框颜色
            HSSFColor leftColor = ((HSSFWorkbook) workbook).getCustomPalette().getColor(((HSSFCellStyle) style).getLeftBorderColor());
            map.put("leftColor", getHssfRgb(leftColor, false));
            // 右边框颜色
            HSSFColor rightColor = ((HSSFWorkbook) workbook).getCustomPalette().getColor(((HSSFCellStyle) style).getRightBorderColor());
            map.put("rightColor", getHssfRgb(rightColor, false));
            // 背景色
            HSSFColor backgroundColor = ((HSSFWorkbook) workbook).getCustomPalette().getColor(((HSSFCellStyle) style).getFillForegroundColor());
            map.put("backgroundColor", getHssfRgb(backgroundColor, true));
        }
//        map.put("borderTop", String.valueOf(style.getBorderTop().getCode() > 1 ? 3 : ((float)style.getBorderTop().getCode() / 2)));
//        map.put("borderBottom", String.valueOf(style.getBorderBottom().getCode() > 1 ? 3 : ((float)style.getBorderBottom().getCode() / 2)));
//        map.put("borderLeft", String.valueOf(style.getBorderLeft().getCode() > 1 ? 3 : ((float)style.getBorderLeft().getCode() / 2)));
//        map.put("borderRight", String.valueOf(style.getBorderRight().getCode() > 1 ? 3 : ((float)style.getBorderRight().getCode() / 2)));
        map.put("borderTop", String.valueOf((float)style.getBorderTop().getCode() / 2));
        map.put("borderBottom", String.valueOf((float)style.getBorderBottom().getCode() / 2));
        map.put("borderLeft", String.valueOf((float)style.getBorderLeft().getCode() / 2));
        map.put("borderRight", String.valueOf((float)style.getBorderRight().getCode() / 2));
    }

    /**
     * Xssf转换Rgb
     *
     * @param xssfColor xssfColor
     * @param flag      标识
     * @return rgb
     */
    private static String getXssfRgb(XSSFColor xssfColor, boolean flag) {
        if (xssfColor != null) {
            byte[] brgb = xssfColor.getRGB();
            if (null != brgb) {
                return "R:" + (brgb[0] & 0xFF) + ",G:" + (brgb[1] & 0xFF) + ",B:" + (brgb[2] & 0xFF);
            }
        }
        if (flag) {
            return "R:255,G:255,B:255";
        } else {
            return "R:0,G:0,B:0";
        }
    }

    /**
     * Hssf转换Rgb
     *
     * @param hssfColor hssfColor
     * @param flag      标识
     * @return rgb
     */
    private static String getHssfRgb(HSSFColor hssfColor, boolean flag) {
        if (hssfColor != null) {
            short[] srgb = hssfColor.getTriplet();
            if (null != srgb) {
                int red = ((byte) srgb[0]) & 0xFF;
                int green = ((byte) srgb[1]) & 0xFF;
                int blue = ((byte) srgb[2]) & 0xFF;
                if (flag) {
                    if (red == 0 && green == 0 & blue == 0) {
                        red = 255;
                        green = 255;
                        blue = 255;
                    }
                }
                return "R:" + red + ",G:" + green + ",B:" + blue;
            }
        }
        if (flag) {
            return "R:255,G:255,B:255";
        } else {
            return "R:0,G:0,B:0";
        }
    }

    /**
     * 挂账结算合并单元格
     *
     * @param fileName 文件路径全量
     * @param firstRow 开始行
     * @param lastRow  结束行 （开始行 + 数据总量）
     * @param lastCol  结束列
     * @throws IOException IO
     */
    public static void mergeCells(String fileName, int firstRow, int lastRow, int lastCol) throws IOException {
        Workbook book = new XSSFWorkbook(new FileInputStream(fileName));
        Sheet sheet = book.getSheetAt(0);
        for (int i = 0; i < lastCol; i++) {
            CellRangeAddress region = new CellRangeAddress(firstRow, lastRow, i, i);
            sheet.addMergedRegion(region);
        }
        File file = new File(fileName);
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        book.write(fileOutputStream);
        fileOutputStream.close();
    }
}
