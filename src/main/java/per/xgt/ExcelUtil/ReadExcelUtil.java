package per.xgt.ExcelUtil;


import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

/**
 * @Author: Valentino
 * @QQ: 3289668817
 * @Email：gentao.xiong
 * @CreateTime: 2020-08-27 10:13:41
 * @Descirption:
 *  excel读写：
 *      method1:importExcelByFilePath(String filePath)读取excel文件，将文件中第一行标题作为key封装到map集合中
 *      method2:importExcelByPath(String filePath)读取excel文件,将文件中列号key(从0开始)封装到map集合中
 */
public class ReadExcelUtil {

    private static Logger logger = LogManager.getLogger(ReadExcelUtil.class);
    private Workbook wb;
    private Sheet sheet;
    private Row row;
    private String[] title;
    @SuppressWarnings("unused")
    private int mapIndex;

    /**
     *读取excel总方法
     * @param filePath
     * @return
     */
    public Map<Integer, Map<String,Object>> importExcelByFilePath(String filePath){
        HashMap<Integer, Map<String, Object>> result = new HashMap<>();
        try {
            initExcel(filePath);
            //获取excel标题
            readExcelTitle();
            //获取excel表格内容
            result = readExcelContent();
        }catch (Exception e){
            logger.error(e.getMessage(), e);
        }
        return result;
    }
    public Map<Integer,Map<Integer,Object>> importExcelByPath(String filePath){
        Map<Integer, Map<Integer, Object>> result = new HashMap<>();
        try{
            initExcel(filePath);
            //获取excel标题
            readExcelTitle();
            //获取excel表格内容
            result = readExcelContent1();
        } catch (Exception e) {
            logger.error(e.getMessage(), e);
        }
        return result;
    }

    //获取excel表格内容
    private Map<Integer, Map<Integer, Object>> readExcelContent1() throws Exception {
        if (wb == null){
            throw new Exception("Workbook对象为空!");
        }
        HashMap<Integer, Map<Integer, Object>> content = new HashMap<>();
        sheet = wb.getSheetAt(0);
        //获取Excel行数
        int rowNum = sheet.getLastRowNum();
        row = sheet.getRow(0);
        int colNum = row.getPhysicalNumberOfCells();
        //正文从第二行开始读取，第一行为列名
        for (int i = 0; i <= rowNum ; i++) {
            row = sheet.getRow(i);
            //row为空继续下一个循环
            if (null == row){
                continue;
            }
            //处理空行
            Iterator<Cell> cellItr = row.cellIterator();
            while (cellItr.hasNext()){
                Cell c = cellItr.next();
                if (c.getCellType() != Cell.CELL_TYPE_BLANK){
                    int j = 0;
                    HashMap<Integer, Object> cellValue = new HashMap<>();
                    while (j < colNum){
                        Object obj = getCellFormatValue(row.getCell(j));
                        cellValue.put(j, obj);
                        j++;
                    }
                    content.put(i, cellValue);
                    //每一行数据只读取一次，这里要停止单元格的单元循环，并进行下一行的数据读取
                    break;
                }
            }
        }
        return content;
    }

    //获取excel表格内容
    private HashMap<Integer, Map<String, Object>> readExcelContent() throws Exception {

        if (wb == null){
            throw new Exception("Workbook对象为空");
        }
        HashMap<Integer, Map<String, Object>> content = new HashMap<Integer, Map<String, Object>>();
        sheet = wb.getSheetAt(0);
        //获取Excel中行数
        int rowNum = sheet.getLastRowNum();
        row = sheet.getRow(0);
        int colNum = row.getPhysicalNumberOfCells();
        //正文从第二行开始读取，第一行为列名
        for (int i = 1; i <= rowNum; i++) {
            row = sheet.getRow(i);
            int j = 0;
            Map<String, Object> cellValue = new HashMap<>();
            while (j < colNum){
                Object obj = getCellFormatValue(row.getCell(j));
                cellValue.put(title[j],obj);
                j++;
            }
            content.put(i, cellValue);
        }
        return content;
    }

    //初始化Excel
    private void initExcel(String filePath){
        if (filePath == null){
            return ;
        }
        String ext = filePath.substring(filePath.lastIndexOf("."));
        try{
            FileInputStream is = new FileInputStream(filePath);
            if (".xls".equals(ext)){
                wb = new HSSFWorkbook(is);
            }else if (".xlsx".equals(ext)){
                wb = new XSSFWorkbook(is);
            }else {
                wb = null;
            }
        } catch (FileNotFoundException e) {
            logger.error(e.getMessage(), e);
        } catch (IOException e) {
            logger.error(e.getMessage(), e);
        }
    }

    //读取Excel的标题行
    private String[] readExcelTitle() throws Exception {
        if (wb == null){
            throw new Exception("Workbook对象为空");
        }
        wb.getSheetAt(0);
        row = sheet.getRow(0);
        int colNum = row.getPhysicalNumberOfCells();
        title = new String[colNum];
        for (int i = 0; i < colNum; i++) {
            Cell cell = row.getCell(i);
            title[i] = (String) getCellFormatValue(cell);
        }
        return title;
    }

    //格式化所属列的单元格格式
    private Object getCellFormatValue(Cell cell){
        Object cellvalue = "";
        if (cell != null){
            //判断当前Cell的type
            switch (cell.getCellType()){
                case Cell.CELL_TYPE_NUMERIC:
                case Cell.CELL_TYPE_FORMULA:{
                    if (DateUtil.isCellDateFormatted(cell)){//date
                        Date date = cell.getDateCellValue();
                        cellvalue = date;
                    }else { //数字
                        cellvalue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                }
                case Cell.CELL_TYPE_STRING:
                    cellvalue = cell.getRichStringCellValue().toString();
                    break;
                default:
                    cellvalue = "";
            }
        }else {
            cellvalue = "";
        }
        return cellvalue;
    }
}
