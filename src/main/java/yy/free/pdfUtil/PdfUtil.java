package yy.free.pdfUtil;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import yy.free.excelUtil.ExcelUtils;
import yy.free.excelUtil.modle.ExcelBaseCell;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * excel转pdf
 * User: yangyong Date:2023/4/616:42 ProjectName: financial-settle Version: 1.1.0
 */
public class PdfUtil {

    /**
     * 字体大小
     */
    private static final float FONT_SIZE = 12;

    /**
     * 大于1MB 用easyExcel解析
     *
     * @param fileName    excel文件名(全路径)
     * @param pdfName     pdf文件名(全路径)
     * @param pdfFontSize pdf字体大小
     * @throws IOException       读写异常
     * @throws DocumentException 转pdf文档异常
     */
    public static void convertExcelToPdf(String fileName, String pdfName, Float pdfFontSize) throws IOException, DocumentException {
        File file = new File(fileName);
        if (file.exists() && file.isFile()) {
            convertNewFileName(fileName, pdfName, pdfFontSize);
        }
    }

    /**
     * excel直接转化pdf
     *
     * @param fileName    excel文件名(全路径)
     * @param pdfName     pdf文件名(全路径)
     * @param pdfFontSize pdf字体大小
     * @throws IOException       读写异常
     * @throws DocumentException 转pdf文档异常
     */
    public static void convertNewFileName(String fileName, String pdfName, Float pdfFontSize) throws DocumentException, IOException {
        // 设置页面大小
        // 定义A3页面大小
        Rectangle rectPageSize = new Rectangle(PageSize.A3);
        // 设置为横版
        rectPageSize = rectPageSize.rotate();
        // 设置边距
        Document document = new Document(rectPageSize, -80, -80, 50, 0);
        Workbook workbook = null;
        if (fileName != null && !"".equals(fileName)) {
            String fileType = fileName.substring(fileName.lastIndexOf(".") + 1);
            workbook = getWorkbook(new FileInputStream(fileName), fileType);
        }
        // 在指定路径创建一个pdf文档
        PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(pdfName));
        if (null != workbook) {
            // 获取Sheet
            Sheet sheet = workbook.getSheetAt(0);
            // 生成PdfTable
            PdfPTable table = getPdfTable(sheet, pdfFontSize);
            document.open();
            document.add(table);
            document.close();
        }
        writer.close();
    }

    /**
     * 获取Workbook
     *
     * @param inputStream 输入流
     * @param fileType    文件类型
     * @return Workbook
     */
    private static Workbook getWorkbook(InputStream inputStream, String fileType) throws IOException {
        Workbook workbook = null;
        if (fileType.equalsIgnoreCase("xls")) {
            workbook = new HSSFWorkbook(inputStream);
        } else if (fileType.equalsIgnoreCase("xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        }
        return workbook;
    }

    /**
     * 获取PdfPTable
     *
     * @param sheet    表单
     * @param fontSize 字体大小
     * @return PdfPTable
     */
    private static PdfPTable getPdfTable(Sheet sheet, Float fontSize) {
        // 获取最大列值
        int column = ExcelUtils.getSheetCellNumn(sheet);
        // 构建table 的整体区域
        PdfPTable table = new PdfPTable(column);
        // 构造
        List<ExcelBaseCell> pdfCells = ExcelUtils.imptBasePdfCell(sheet, column);
        for (ExcelBaseCell basePdfCell : pdfCells) {
            table.addCell(mergeColAndRow(basePdfCell, fontSize));
        }
        return table;
    }

    public static PdfPCell mergeColAndRow(ExcelBaseCell basePdfCell, Float fontSize) {
        PdfPCell cell = null;

        switch (basePdfCell.getCellType()) {
            case NUMERIC:
                cell = mergeCellNew(basePdfCell.getValue().toString(), basePdfCell.getRowSpan(),
                        basePdfCell.getColumnSpan(), basePdfCell.getStlys(), fontSize);
                break;
            case STRING:
                cell = mergeCellNew(basePdfCell == null ? "" : (String) basePdfCell.getValue(),
                        basePdfCell.getRowSpan(), basePdfCell.getColumnSpan(), basePdfCell.getStlys(), fontSize);
                break;
            case FORMULA:
                cell = mergeCellNew(String.valueOf(basePdfCell.getValue()), basePdfCell.getRowSpan(),
                        basePdfCell.getColumnSpan(), basePdfCell.getStlys(), fontSize);
                break;

            case BOOLEAN:
                cell = mergeCellNew((String) basePdfCell.getValue(), basePdfCell.getRowSpan(),
                        basePdfCell.getColumnSpan(), basePdfCell.getStlys(), fontSize);
                ;
                break;
            case BLANK:
            default:
                cell = mergeCellNew("", basePdfCell.getRowSpan(), basePdfCell.getColumnSpan(), basePdfCell.getStlys(), fontSize);
                break;
        }
        return cell;
    }

    /**
     * 处理单元格
     *
     * @param str      值
     * @param i        行
     * @param j        列
     * @param stlys    样式
     * @param fontSize 字体大小
     * @return PdfPCell
     */
    private static PdfPCell mergeCellNew(String str, int i, int j, Map<String, String> stlys, Float fontSize) {
        //创建BaseFont对象，指明字体，编码方式,是否嵌入
        BaseFont bf = null;
        try {
            bf = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
        } catch (DocumentException | IOException e) {
            e.printStackTrace();
        }
        if (fontSize == null) {
            fontSize = FONT_SIZE;
        }
        //创建Font对象，将基础字体对象，字体大小，字体风格
        Font font = new Font(bf, fontSize, Font.NORMAL);
        // 上边框颜色，默认黑色
        BaseColor topBorderColor = getBaseColor(stlys.get("topBorderColor"));
        // 下边框颜色，默认黑色
        BaseColor bottomBorderColor = getBaseColor(stlys.get("bottomColor"));
        // 左边框颜色，默认黑色
        BaseColor leftBorderColor = getBaseColor(stlys.get("leftColor"));
        // 右边框颜色，默认黑色
        BaseColor rightBorderColor = getBaseColor(stlys.get("rightColor"));
        // 背景色，默认白色
        BaseColor backgroundColor = getBaseColor(stlys.get("backgroundColor"));
        if (!"".equals(str)) {
            //获取字体颜色
            int red = Integer.parseInt(stlys.get("font_red"));
            int green = Integer.parseInt(stlys.get("font_green"));
            int blue = Integer.parseInt(stlys.get("font_blue"));
            font.setStyle(Integer.parseInt(stlys.get("bold")));
            font.setColor(red, green, blue);
        }
        PdfPCell cell = new PdfPCell(new Paragraph(str, font));
        cell.setUseVariableBorders(true);
        cell.setMinimumHeight(12);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        cell.setRowspan(i);
        cell.setColspan(j);
        // 设置单元格背景色
        cell.setBackgroundColor(backgroundColor);
        // 渲染pdf单元格边框颜色
        cell.setBorderColorTop(topBorderColor);
        cell.setBorderColorBottom(bottomBorderColor);
        cell.setBorderColorLeft(leftBorderColor);
        cell.setBorderColorRight(rightBorderColor);
        // 设置单元格边框样式
        cell.setBorderWidthTop(Float.parseFloat(stlys.get("borderTop")));
        cell.setBorderWidthBottom(Float.parseFloat(stlys.get("borderBottom")));
        cell.setBorderWidthLeft(Float.parseFloat(stlys.get("borderLeft")));
        cell.setBorderWidthRight(Float.parseFloat(stlys.get("borderRight")));
        return cell;
    }

    /**
     * 获取RGB颜色
     *
     * @param rgb rgb
     * @return BaseColor
     */
    private static BaseColor getBaseColor(String rgb) {
        String[] rgbList = rgb.split(",");
        Map<String, String> rgbMap = new HashMap<>();
        for (String test1 : rgbList) {
            rgbMap.put(test1.split(":")[0], test1.split(":")[1]);
        }
        return new BaseColor(Integer.parseInt(rgbMap.get("R")),
                Integer.parseInt(rgbMap.get("G")), Integer.parseInt(rgbMap.get("B")));
    }
}
