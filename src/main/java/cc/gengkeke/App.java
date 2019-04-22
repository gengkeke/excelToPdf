package cc.gengkeke;

import com.itextpdf.text.DocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class App {
    public static void main(String[] args) {
        try {
           /* File file = new File("E:\\excel2img\\天工大数据-20181126105446622.xlsx");
            excel2Pdf(file, "天工大数据-20181126105446622", "E:\\excel2img\\");*/
            File file = new File("E:\\excel2img\\天工大数据-test3.xlsx");
            excel2Pdf(file, "天工大数据-test3", "E:\\excel2img\\");
        } catch (IOException e) {
            e.printStackTrace();
        } catch (DocumentException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    private static void excel2Pdf(File file, String excelFileName, String pdfOutDir) throws IOException, InvalidFormatException, DocumentException {
        FileInputStream inputStream = null;
        FileOutputStream outputStream = null;
        try {
            inputStream = new FileInputStream(file);
            Workbook workbook = WorkbookFactory.create(inputStream);
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName().replace(" ", "");
                outputStream = new FileOutputStream(new File(pdfOutDir + excelFileName + "_" + sheetName + ".pdf"));
                Excel2Pdf.convert(sheetName, sheet, outputStream, workbook);
            }
        } finally {
            if (outputStream != null) {
                outputStream.close();
            }
            if (inputStream != null) {
                inputStream.close();
            }
        }
    }
}
