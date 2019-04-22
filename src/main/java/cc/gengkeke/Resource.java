package cc.gengkeke;

import com.itextpdf.text.pdf.BaseFont;

/**
 * @Author: 18073004
 * @Date: 2019/1/2 09:03
 * @Description:
 */
public class Resource {
    /**
     * 中文字体支持
     */
    protected static BaseFont BASE_FONT_CHINESE;
    static {
        try {
            BASE_FONT_CHINESE = BaseFont.createFont("STSongStd-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}