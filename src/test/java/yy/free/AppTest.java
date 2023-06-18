package yy.free;

import com.itextpdf.text.DocumentException;
import org.junit.Test;
import yy.free.pdfUtil.PdfUtil;

import java.io.IOException;

/**
 * Unit test for simple App.
 */
public class AppTest 
{
    /**
     * Rigorous Test :-)
     */
    @Test
    public void excelConvertPdf() throws DocumentException, IOException {
        // excel绝对路径
        String excelPath = "D:\\3pdf.xlsx";
//        String excelPath = "D:\\3pdf.xls";
        // pdf绝对路径
        String pdfPath = "D:\\3pdf.pdf";
        // pdf字体大小
        Float pdfFontSize = 5f;
        PdfUtil.convertExcelToPdf(excelPath, pdfPath, pdfFontSize);
    }
}
