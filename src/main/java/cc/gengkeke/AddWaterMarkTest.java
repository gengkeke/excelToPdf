package cc.gengkeke;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Element;
import com.itextpdf.text.pdf.*;

import java.io.File;
import java.io.FileOutputStream;

/**
 * @Author: 18073004
 * @Date: 2019/1/3 14:22
 * @Description:
 */
public class AddWaterMarkTest {
    public static void main(String[] args) {
        try {
            addWaterMark();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void addWaterMark() throws Exception {
        String srcFile = "E:\\excel2img\\天工大数据-20181126105446622.pdf";//要添加水印的文件
        String text = "苏宁沙箱系统      耿可可-18073004";//要添加水印的内容
        PdfReader reader = new PdfReader(srcFile);// 待加水印的文件
        PdfStamper stamper = new PdfStamper(reader, new FileOutputStream(new File("E:\\excel2img\\天工大数据-20181126105446622-1.pdf")));// 加完水印的文件
//          byte[] userPassword = "123".getBytes();
        byte[] ownerPassword = "sandbox".getBytes();
//          int permissions = PdfWriter.ALLOW_COPY|PdfWriter.ALLOW_MODIFY_CONTENTS|PdfWriter.ALLOW_PRINTING;
//          stamper.setEncryption(null, ownerPassword, permissions,false);
        stamper.setEncryption(null, ownerPassword, PdfWriter.ALLOW_ASSEMBLY, false);
        stamper.setEncryption(null, ownerPassword, PdfWriter.ALLOW_COPY, false);
        stamper.setEncryption(null, ownerPassword, PdfWriter.ALLOW_DEGRADED_PRINTING, false);
        stamper.setEncryption(null, ownerPassword, PdfWriter.ALLOW_FILL_IN, false);
        stamper.setEncryption(null, ownerPassword, PdfWriter.ALLOW_MODIFY_ANNOTATIONS, false);
        stamper.setEncryption(null, ownerPassword, PdfWriter.ALLOW_MODIFY_CONTENTS, false);
        stamper.setEncryption(null, ownerPassword, PdfWriter.ALLOW_PRINTING, false);
        stamper.setEncryption(null, ownerPassword, PdfWriter.ALLOW_SCREENREADERS, false);
        stamper.setEncryption(null, ownerPassword, PdfWriter.DO_NOT_ENCRYPT_METADATA, false);
        stamper.setViewerPreferences(PdfWriter.HideToolbar | PdfWriter.HideMenubar);
        //stamper.setViewerPreferences(PdfWriter.HideWindowUI);//这句话的注释打开，在IE8也能使用
        int total = reader.getNumberOfPages() + 1;
        PdfContentByte content;
        //BaseFont font = BaseFont.createFont("font/SIMKAI.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
        // 循环对每页插入水印
        PdfGState gs = new PdfGState();
        gs.setFillOpacity(0.15f);
        gs.setStrokeOpacity(0.15f);
        for (int i = 1; i < total; i++) {
            content = stamper.getOverContent(i);//水印在之前文本上
            //content = stamper.getUnderContent(i);// 水印在之前文本下
            content.beginText();// 开始
            content.setColorFill(new BaseColor(119, 136, 153));// 设置颜色
            content.setFontAndSize(Resource.BASE_FONT_CHINESE, 10);// 设置字体及字号
            content.setTextMatrix(0, 0);// 设置起始位置
            content.setGState(gs);//水印 透明度
            for (int j = 0; j < 10; j++) {
                for (int k = 0; k <10 ; k++) {
                    content.showTextAligned(Element.ALIGN_LEFT, text, k*180, j*90, 15);//开始写入水印
                }
            }
            content.endText();
            //content.setGState(gs);// 图片水印 透明度
            //content.addImage(img);// 图片水印
        }
        stamper.close();
        reader.close();
    }
}
