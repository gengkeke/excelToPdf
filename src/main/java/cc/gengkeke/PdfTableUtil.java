package cc.gengkeke;

import com.itextpdf.text.*;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.io.IOException;
import java.net.MalformedURLException;
import java.util.ArrayList;
import java.util.List;

/**
 * @Author: 18073004
 * @Date: 2019/1/2 09:03
 * @Description: PDFTable
 */
public class PdfTableUtil {
    private static boolean setting = false;

    public static PdfPTable toParseContent(Sheet sheet, Workbook workbook, String sheetName) throws BadElementException, IOException {
        int lastRowNum = sheet.getLastRowNum();
        List<PdfPCell> cells = new ArrayList<>();
        float[] widths = null;
        float mw = 0;
        for (int i = 0; i < lastRowNum; i++) {
            Row row = sheet.getRow(i);
            int maxCol = getMaxCol(sheet);
            float[] cws = new float[maxCol];
            for (int j = 0; j < maxCol; j++) {
                Cell cell = row.getCell(j);
                if (cell == null) {
                    cell = row.createCell(j);
                    cell.setCellStyle(row.getCell(0).getCellStyle());
                }
                float cw = getPOIColumnWidth(cell, sheet);
                cws[j] = cw;
                if (isUsed(j, row.getRowNum(), sheet)) {
                    continue;
                }
                cell.setCellType(Cell.CELL_TYPE_STRING);
                CellRangeAddress range = getColspanRowspanByExcel(row.getRowNum(), cell.getColumnIndex(), sheet);
                int rowSpan = 1;
                int colSpan = 1;
                if (range != null) {
                    rowSpan = range.getLastRow() - range.getFirstRow() + 1;
                    colSpan = range.getLastColumn() - range.getFirstColumn() + 1;
                }
                //PDF单元格
                PdfPCell pdfpCell = new PdfPCell();
                pdfpCell.setBackgroundColor(new BaseColor(getBackgroundColorByExcel(cell.getCellStyle())));
                pdfpCell.setColspan(colSpan);
                pdfpCell.setRowspan(rowSpan);
                pdfpCell.setVerticalAlignment(getVAlignByExcel(cell.getCellStyle().getVerticalAlignment()));
                pdfpCell.setHorizontalAlignment(getHAlignByExcel(cell.getCellStyle().getAlignment()));
                pdfpCell.setPhrase(getPhrase(cell, sheetName, workbook));
                pdfpCell.setFixedHeight(getPixelHeight(row.getHeightInPoints()));
                addBorderByExcel(pdfpCell, cell.getCellStyle(), workbook);
                //addImageByPOICell(pdfpCell, cell, cw);
                cells.add(pdfpCell);
                j += colSpan - 1;
            }
            float rw = 0;
            for (float c : cws) {
                rw += c;
            }
            if (rw > mw || mw == 0) {
                widths = cws;
                mw = rw;
            }
        }
        if (widths != null) {
            PdfPTable table = new PdfPTable(widths);
            table.setWidthPercentage(100);
            //table.setLockedWidth(true);
            for (PdfPCell pdfpCell : cells) {
                table.addCell(pdfpCell);
            }
            return table;
        }
        return null;
    }

    private static int getMaxCol(Sheet sheet) {
        int n = 1;
        int rowLength = sheet.getLastRowNum();
        for (int i = 0; i < rowLength; i++) {
            Row row = sheet.getRow(i);
            short lastCellNum = row.getLastCellNum();
            if (lastCellNum > n) {
                n = lastCellNum;
            }

        }
        return n;
    }

    private static Phrase getPhrase(Cell cell, String sheetName, Workbook workbook) {
        if (setting || sheetName == null) {
            return new Phrase(cell.getStringCellValue(), getFontByExcel(cell.getCellStyle(), workbook));
        }
        Anchor anchor = new Anchor(cell.getStringCellValue(), getFontByExcel(cell.getCellStyle(), workbook));
        anchor.setName(sheetName);
        setting = true;
        return anchor;
    }


    private static void addImageByPOICell(PdfPCell pdfpCell, Cell cell, float cellWidth) throws BadElementException, MalformedURLException, IOException {
        POIImage poiImage = new POIImage().getCellImage(cell);
        byte[] bytes = poiImage.getBytes();
        if (bytes != null) {
//           double cw = cellWidth;
//           double ch = pdfpCell.getFixedHeight();
//           double iw = poiImage.getDimension().getWidth();
//           double ih = poiImage.getDimension().getHeight();
//           double scale = cw / ch;
//           double nw = iw * scale;
//           double nh = ih - (iw - nw);
//           POIUtil.scale(bytes , nw  , nh);
            pdfpCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            pdfpCell.setHorizontalAlignment(Element.ALIGN_CENTER);
            Image image = Image.getInstance(bytes);
            pdfpCell.setImage(image);
        }
    }

    private static float getPixelHeight(float poiHeight) {
        return poiHeight / 28.6f * 26f;
    }

    /**
     * <p>Description: 此处获取Excel的列宽像素(无法精确实现,期待有能力的朋友进行改善此处)</p>
     *
     * @param cell
     * @return 像素宽
     */
    private static int getPOIColumnWidth(Cell cell, Sheet sheet) {
        if (cell == null) {
            return 416;
        }
        int colWidthpoi = sheet.getColumnWidth(cell.getColumnIndex());
        int widthPixel = 0;
        if (colWidthpoi >= 416) {
            widthPixel = (int) (((colWidthpoi - 416.0) / 256.0) * 8.0 + 13.0 + 0.5);
        } else {
            widthPixel = (int) (colWidthpoi / 416.0 * 13.0 + 0.5);
        }
        return widthPixel;
    }

    private static CellRangeAddress getColspanRowspanByExcel(int rowIndex, int colIndex, Sheet sheet) {
        CellRangeAddress result = null;
        int num = sheet.getNumMergedRegions();
        for (int i = 0; i < num; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            if (range.getFirstColumn() == colIndex && range.getFirstRow() == rowIndex) {
                result = range;
            }
        }
        return result;
    }

    private static boolean isUsed(int colIndex, int rowIndex, Sheet sheet) {
        boolean result = false;
        int num = sheet.getNumMergedRegions();
        for (int i = 0; i < num; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            if (firstRow < rowIndex && lastRow >= rowIndex) {
                if (firstColumn <= colIndex && lastColumn >= colIndex) {
                    result = true;
                }
            }
        }
        return result;
    }

    private static Font getFontByExcel(CellStyle style, Workbook workbook) {
        Font result = new Font(Resource.BASE_FONT_CHINESE, 8, Font.NORMAL);
        //字体样式索引
        org.apache.poi.ss.usermodel.Font font = workbook.getFontAt(style.getFontIndex());
        int colorIndex = font.getColor();
        if (font.getBoldweight() == org.apache.poi.ss.usermodel.Font.BOLDWEIGHT_BOLD) {
            result.setStyle(Font.BOLD);
        }
        if (font instanceof HSSFFont) {
            HSSFColor color = HSSFColor.getIndexHash().get(colorIndex);
            if (color != null) {
                int rbg = POIUtil.getRGB(color);
                result.setColor(new BaseColor(rbg));
            }
        }

        if (font instanceof XSSFFont) {
            XSSFColor color = ((XSSFFont) font).getXSSFColor();
            if (color != null) {
                int rbg = POIUtil.getRGB(color);
                // result.setColor(new BaseColor(rbg));
            }
        }
       /* if (color != null) {
            int rbg = POIUtil.getRGB(color);
            result.setColor(new BaseColor(rbg));
        }*/
        //下划线
        FontUnderline underline = FontUnderline.valueOf(font.getUnderline());
        if (underline == FontUnderline.SINGLE) {
            String ulString = Font.FontStyle.UNDERLINE.getValue();
            result.setStyle(ulString);
        }
        return result;
    }

    private static int getBackgroundColorByExcel(CellStyle style) {
        Color color = style.getFillForegroundColorColor();
        return POIUtil.getRGB(color);
    }

    private static void addBorderByExcel(PdfPCell cell, CellStyle style, Workbook workbook) {
        cell.setBorderColorLeft(new BaseColor(POIUtil.getBorderRBG(workbook, style.getLeftBorderColor(), style)));
        cell.setBorderColorRight(new BaseColor(POIUtil.getBorderRBG(workbook, style.getRightBorderColor(), style)));
        cell.setBorderColorTop(new BaseColor(POIUtil.getBorderRBG(workbook, style.getTopBorderColor(), style)));
        cell.setBorderColorBottom(new BaseColor(POIUtil.getBorderRBG(workbook, style.getBottomBorderColor(), style)));
    }

    private static int getVAlignByExcel(short align) {
        int result = 0;
        if (align == CellStyle.VERTICAL_BOTTOM) {
            result = Element.ALIGN_BOTTOM;
        }
        if (align == CellStyle.VERTICAL_CENTER) {
            result = Element.ALIGN_MIDDLE;
        }
        if (align == CellStyle.VERTICAL_JUSTIFY) {
            result = Element.ALIGN_JUSTIFIED;
        }
        if (align == CellStyle.VERTICAL_TOP) {
            result = Element.ALIGN_TOP;
        }
        return result;
    }

    private static int getHAlignByExcel(short align) {
        int result = 0;
        if (align == CellStyle.ALIGN_LEFT) {
            result = Element.ALIGN_LEFT;
        }
        if (align == CellStyle.ALIGN_RIGHT) {
            result = Element.ALIGN_RIGHT;
        }
        if (align == CellStyle.ALIGN_JUSTIFY) {
            result = Element.ALIGN_JUSTIFIED;
        }
        if (align == CellStyle.ALIGN_CENTER) {
            result = Element.ALIGN_CENTER;
        }
        return result;
    }
}
