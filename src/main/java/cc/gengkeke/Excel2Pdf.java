package cc.gengkeke;

import com.itextpdf.text.*;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.IOException;
import java.io.OutputStream;
import java.net.MalformedURLException;
import java.util.ArrayList;
import java.util.List;

/**
 * @Author: 18073004
 * @Date: 2019/1/2 09:03
 * @Description:
 */
public class Excel2Pdf {
    /**
     * 转换调用
     *
     * @throws DocumentException
     * @throws MalformedURLException
     * @throws IOException
     */
    public static void convert(String sheetName, Sheet sheet, OutputStream outputStream, Workbook workbook) throws DocumentException, IOException{
        //Rectangle rect = new Rectangle(PageSize.A0.rotate());
        Rectangle rect = new Rectangle(new RectangleReadOnly(24400, 2384));
        //rect.setBackgroundColor(BaseColor.ORANGE);
        Document document = new Document(rect);
        //边距
        document.setMargins(10, 10, 10, 10);
        PdfWriter writer = PdfWriter.getInstance(document, outputStream);
        //设置权限
        writer.setEncryption(null, null, PdfWriter.ALLOW_COPY, false);
        writer.setEncryption(null, null, PdfWriter.ALLOW_ASSEMBLY, false);
        writer.setEncryption(null, null, PdfWriter.ALLOW_COPY, false);
        writer.setEncryption(null, null, PdfWriter.ALLOW_DEGRADED_PRINTING, false);
        writer.setEncryption(null, null, PdfWriter.ALLOW_FILL_IN, false);
        writer.setEncryption(null, null, PdfWriter.ALLOW_MODIFY_ANNOTATIONS, false);
        writer.setEncryption(null, null, PdfWriter.ALLOW_MODIFY_CONTENTS, false);
        writer.setEncryption(null, null, PdfWriter.ALLOW_PRINTING, false);
        writer.setEncryption(null, null, PdfWriter.ALLOW_SCREENREADERS, false);
        writer.setEncryption(null, null, PdfWriter.DO_NOT_ENCRYPT_METADATA, false);
        writer.setPageEvent(new PDFPageEvent());
        document.open();
        // 内容索引创建
        toCreateContentIndexes(document, sheetName, writer);
        PdfPTable table = toCreatePdfTable(sheet, workbook, sheetName);
        if (table != null) {
            document.add(table);
        }
        List<byte[]> bytes = toGetImageBytes(sheet);
        if (!bytes.isEmpty()) {
            for (byte[] data : bytes) {
                document.add(Image.getInstance(data));
            }
        }
        document.close();
        writer.close();
    }

    private static List<byte[]> toGetImageBytes(Sheet sheet) throws IOException, BadElementException {
        List<byte[]> datas = new ArrayList<>();
        if (sheet instanceof XSSFSheet) {
            XSSFDrawing xssfDrawing = (XSSFDrawing) sheet.getDrawingPatriarch();
            if (xssfDrawing != null) {
                List<XSSFShape> shapes = xssfDrawing.getShapes();
                if (!shapes.isEmpty()) {
                    for (Shape shape : shapes) {
                        if (shape instanceof XSSFPicture) {
                            XSSFPicture picture = (XSSFPicture) shape;
                            byte[] data = picture.getPictureData().getData();
                            datas.add(data);
                        }
           /* if (shape instanceof XSSFGraphicFrame) {
                XSSFGraphicFrame xssfGraphicFrame = (XSSFGraphicFrame) shape;
                // xssfGraphicFrame.getDrawing()
                CTGraphicalObjectData graphicData = xssfGraphicFrame.getCTGraphicalObjectFrame().getGraphic().getGraphicData();
                graphicData.save(new File("E:\\excel2img\\test.png"));
            }*/
                    }
                }
            }
        }
        //嵌入文件
        // workbook.getAllEmbedds();
        return datas;
    }

    private static PdfPTable toCreatePdfTable(Sheet sheet, Workbook workbook, String sheetName) throws IOException, DocumentException {
        PdfPTable table = PdfTableUtil.toParseContent(sheet, workbook, sheetName);
        if (table != null) {
            table.setKeepTogether(true);
            //table.setWidthPercentage(new float[]{100} , writer.getPageSize());
            table.getDefaultCell().setBorder(PdfPCell.NO_BORDER);
            return table;
        }
        return null;
    }

    /**
     * 内容索引创建
     */
    private static void toCreateContentIndexes(Document document, String sheetName, PdfWriter writer) throws DocumentException {
        PdfPTable table = new PdfPTable(1);
        table.setKeepTogether(true);
        table.getDefaultCell().setBorder(PdfPCell.NO_BORDER);
        Font font = new Font(Resource.BASE_FONT_CHINESE, 12, Font.NORMAL);
        font.setColor(new BaseColor(0, 0, 255));
        Anchor anchor = new Anchor(sheetName, font);
        anchor.setReference("#" + sheetName);
        PdfPCell cell = new PdfPCell(anchor);
        cell.setBorder(0);
        table.addCell(cell);
        document.add(table);
    }

    /**
     * ClassName: PDFPageEvent
     * Description: 事件 -> 页码控制/水印添加
     */
    private static class PDFPageEvent extends PdfPageEventHelper {
        private PdfTemplate template;
        private BaseFont baseFont;

        @Override
        public void onStartPage(PdfWriter writer, Document document) {
            try {
                this.template = writer.getDirectContent().createTemplate(100, 100);
                this.baseFont = new Font(Resource.BASE_FONT_CHINESE, 8, Font.NORMAL).getBaseFont();
            } catch (Exception e) {
                throw new ExceptionConverter(e);
            }
        }

        @Override
        public void onEndPage(PdfWriter writer, Document document) {
            //在每页结束的时候把“第x页”信息写道模版指定位置
            PdfContentByte byteContent = writer.getDirectContent();
            String text = "第" + writer.getPageNumber() + "页";
            float textWidth = this.baseFont.getWidthPoint(text, 8);
            float realWidth = document.right() - textWidth;

            byteContent.beginText();
            byteContent.setFontAndSize(this.baseFont, 10);
            byteContent.setTextMatrix(realWidth, document.bottom());
            byteContent.showText(text);
            byteContent.endText();
            //添加水印//
            byteContent.beginText();
            // 设置颜色
            byteContent.setColorFill(new BaseColor(119, 136, 153));
            // 设置字体及字号
            byteContent.setFontAndSize(Resource.BASE_FONT_CHINESE, 10);
            // 设置起始位置
            byteContent.setTextMatrix(0, 0);
            //水印 透明度
            PdfGState gs = new PdfGState();
            gs.setFillOpacity(0.15f);
            gs.setStrokeOpacity(0.15f);
            byteContent.setGState(gs);
            for (int j = 0; j < 10; j++) {
                for (int k = 0; k < 10; k++) {
                    //开始写入水印
                    byteContent.showTextAligned(Element.ALIGN_LEFT, "苏宁沙箱系统      耿可可-18073004", k * 180, j * 90, 15);
                }
            }
            byteContent.endText();
            byteContent.stroke();
            byteContent.addTemplate(this.template, realWidth, document.bottom());
        }
    }
}
