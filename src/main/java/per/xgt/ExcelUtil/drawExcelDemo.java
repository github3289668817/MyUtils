package per.xgt.ExcelUtil;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @Author: Valentino
 * @QQ: 3289668817
 * @Email：gentao.xiong
 * @CreateTime: 2020-08-27 10:09:12
 * @Descirption: 测试Excel单元格合并Demo
 */
public class drawExcelDemo {

    public static void main(String[] args) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        style.setWrapText(true);//自动换行
        HSSFDataFormat format = workbook.createDataFormat();
        style.setDataFormat(format.getFormat("@"));
        HSSFSheet sheet = workbook.createSheet("sheet");
        for (int i = 0; i < 17; i++) {
            sheet.setColumnWidth(i, 256*20);
        }
        for (int i = 0; i < 17; i++) {
            sheet.setDefaultColumnStyle(i, style);
        }
        //第一行 标题行
        HSSFRow row0 = sheet.createRow(0);
        String[] strRow1 = {"列名","列名","列名","列名","列名","列名","列名","列名","列名","列名","列名","列名","列名","","","","列名"};
        for (int i = 0; i < 17; i++) {
            HSSFCell cell = row0.createCell(i);
            cell.setCellStyle(style);
            cell.setCellValue(strRow1[i]);
        }
        //第二行
        HSSFRow row1 = sheet.createRow(1);
        String[] strRow2 = {"","","","","","","","","","","","","列名","列名","列名","列名",""};
        for (int i = 0; i < 17; i++) {
            HSSFCell cell = row1.createCell(i);
            cell.setCellStyle(style);
            cell.setCellValue(strRow2[i]);
        }
        //合并单元格
        for (int i = 0; i < 12; i++) {
            CellRangeAddress region = new CellRangeAddress(0, 1, i, i);
            sheet.addMergedRegion(region);
        }
        CellRangeAddress region13 = new CellRangeAddress(0, 1, 16, 16);
        CellRangeAddress region14 = new CellRangeAddress(0, 0, 12, 15);
        sheet.addMergedRegion(region13);
        sheet.addMergedRegion(region14);
        File file = new File("E:\\demo.xls");
        FileOutputStream fout = new FileOutputStream(file);
        workbook.write(fout);
        fout.close();
    }
}
