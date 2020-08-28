package per.xgt.ExcelUtil;

import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.joda.time.DateTime;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @Author: Valentino
 * @QQ: 3289668817
 * @Email：gentao.xiong
 * @CreateTime: 2020-08-27 14:19:37
 * @Descirption: 将java集合内数据以Excel的形式输出到IO设备中
 */
public class ExportExcelUtil<T> {

    private static Logger logger = LogManager.getLogger(ExportExcelUtil.class);

    public void exportExcel(Collection<T> dataset, OutputStream out){
        exportExcel("导出Excel", null,dataset,out,"yyyy-MM-dd");
    }
    public void exportExcel(String[] headers,Collection<T> dataset,OutputStream out){
        exportExcel("导出Excel",headers,dataset,out,"yyyy-MM-dd");
    }
    public void exportExcel(String[] headers,Collection<T> dateset,OutputStream out,String pattern){
        exportExcel("导出Excel",headers,dateset,out,pattern);
    }
    public void exportExcel(String[] headers,String[] propertytys,Collection<T> dateset,OutputStream out){
        exportExcel("导出Excel",headers,propertytys,dateset,out,"yyyy-MM-dd");
    }
    public void exportExcel(String title,String[] headers,String[] propertytys,Collection<T> dateset,OutputStream out){
        exportExcel(title,headers,propertytys,dateset,out,"yyyy-MM-dd");
    }

    /**
     * 通用方法，通过反射，将java集合中并符合条件的数据以Excel的形式输入到指定的IO设备中
     * @param title 表格标题名
     * @param headers 表格属性列名数组
     * @param dataset 需要显示的数据集合，且对象符合javabean风格，对象属性数据类型支持基本数据类型及String,Date,byte[]（图）
     * @param out 与输出设备关联的流对象，可以将Excel文档导出到本地文件或者网络中
     * @param pattern 如果有时间数据，设定输出格式，默认为（"yyyy-MM-dd"）
     */
    @SuppressWarnings({"unchecked","deprecation"})
    public void exportExcel(String title,String[] headers,Collection<T> dataset,OutputStream out,String pattern){
        //声明一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //生成一个表格
        HSSFSheet sheet = workbook.createSheet(title);
        //设置表格默认列宽为个字节
        sheet.setDefaultColumnWidth((short)15);
        //生成一个样式
        HSSFCellStyle style = workbook.createCellStyle();
        //设置这些样式
        style.setFillBackgroundColor(HSSFColor.SKY_BLUE.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        //生成一个字体
        HSSFFont font = workbook.createFont();
        font.setColor(HSSFColor.VIOLET.index);
        font.setFontHeightInPoints((short)12);
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        //把字体应用到当前的样式
        style.setFont(font);
        //生成并设置另一个样式
        HSSFCellStyle style2 = workbook.createCellStyle();
        style2.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
        style2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style2.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style2.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style2.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style2.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        //生成另一个字体
        HSSFFont font2 = workbook.createFont();
        font2.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
        //把字体应用到当前的样式
        style2.setFont(font2);
        //声明一个画图的顶级管理器
        HSSFPatriarch patriarch = sheet.createDrawingPatriarch();
        //定义注释的大小和位置
        HSSFComment comment = patriarch.createComment(new HSSFClientAnchor(0, 0, 0, 0, (short) 4, 2, (short) 6, 5));
        //设置注释内容
        comment.setString(new HSSFRichTextString("可以添加注释!"));
        //设置注释作者，当鼠标移动到单元格上是可以在状态栏中看到该内容。
        comment.setAuthor("XGT");
        //产生表格标题行
        HSSFRow row = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            HSSFCell cell = row.createCell(i);
            cell.setCellStyle(style);
            HSSFRichTextString text = new HSSFRichTextString(headers[i]);
            cell.setCellValue(title);

        }
        //遍历集合数据，产生数据行
        Iterator<T> it = dataset.iterator();
        int index = 0;
        while (it.hasNext()){
            index++;
            row = sheet.createRow(index);
            T t = (T) it.next();
            //利用反射，根据javabean属性的先后顺序，动态调用getter方法得到属性值
            Field[] fields = t.getClass().getDeclaredFields();
            for (int i = 0; i < fields.length; i++) {
                HSSFCell cell = row.createCell(i);
                cell.setCellStyle(style2);
                Field field = fields[i];
                String fieldName = field.getName();
                String getMethodName = "get" + fieldName.substring(0,1).toUpperCase() + fieldName.substring(1);
                try {
                    @SuppressWarnings("rawtypes")
                    Class tCls = t.getClass();
                    Method getMethod = tCls.getMethod(getMethodName, new Class[] {});
                    Object value = getMethod.invoke(t, new Object(){});
                    //判断值得类型后进行强制类型转换
                    String textValue = null;
                    if (value instanceof Boolean){
                        @SuppressWarnings("unused")
                        Boolean bValue =(Boolean)value;
                        textValue = value.toString();
                    }else if (value instanceof Date){
                        Date date = (Date) value;
                        SimpleDateFormat sdf = new SimpleDateFormat(pattern);
                        textValue = sdf.format(date);
                    }else if (value instanceof byte[]){
                        //有图片时，设置行高为60px
                        row.setHeightInPoints(60);
                        //设置图片所在列宽高为80px，注意单位换算
                        sheet.setColumnWidth(i, (short)(35.7*80));
                        byte[] bsValue = (byte[]) value;
                        HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 1023, 255, (short) 6, index, (short) 6, index);
                        anchor.setAnchorType(2);
                        patriarch.createPicture(anchor,workbook.addPicture(bsValue, HSSFWorkbook.PICTURE_TYPE_JPEG));
                    }else {
                        //其他数据类型都当作字符串简单处理
                        textValue = value == null ? "" : value.toString();
                    }
                    //如果不是图片数据，就利用正则表达式判断textValue是否全部由数字组成
                    if (textValue != null){
                        Pattern p = Pattern.compile("^//d+(//.//d+)?$");
                        Matcher matcher = p.matcher(textValue);
                        if (matcher.matches()){
                            //是数字当作double处理
                            cell.setCellValue(Double.parseDouble(textValue));
                        }else {
                            HSSFRichTextString richString = new HSSFRichTextString(textValue);
                            HSSFFont font3 = workbook.createFont();
                            font3.setColor(HSSFColor.BLUE.index);
                            richString.applyFont(font3);
                            cell.setCellValue(richString);
                        }
                    }
                } catch (NoSuchMethodException e) {
                    e.printStackTrace();
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                } catch (InvocationTargetException e) {
                    e.printStackTrace();
                }finally {
                    //清空资源
                }
            }
        }
        try {
            workbook.write(out);
        } catch (IOException e) {
            logger.info("{}",e);
        }
    }

    /**
     * 通用方法，通过反射，将java集合中并符合条件的数据以Excel的形式输入到指定的IO设备中
     * @param title 表格标题名
     * @param headers 表格属性列名数组
     * @param propertys 表格属性列对应得对象属性数组
     * @param dataset 需要显示的数据集合，且对象符合javabean风格，对象属性数据类型支持基本数据类型及String,Date,byte[]（图）
     * @param out 与输出设备关联的流对象，可以将Excel文档导出到本地文件或者网络中
     * @param pattern 如果有时间数据，设定输出格式，默认为（"yyyy-MM-dd"）
     */
    @SuppressWarnings({"unchecked","deprecation"})
    public void exportExcel(String title,String[] headers,String[] propertys,Collection<T> dataset,OutputStream out,String pattern){
        //声明一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //生成一个表格
        HSSFSheet sheet = workbook.createSheet(title);
        //设置表格默认列宽为个字节
        sheet.setDefaultColumnWidth((short)15);
        //声明一个画图的顶级管理器
        HSSFPatriarch patriarch = sheet.createDrawingPatriarch();
        //定义注释的大小和位置
        HSSFComment comment = patriarch.createComment(new HSSFClientAnchor(0, 0, 0, 0, (short) 4, 2, (short) 6, 5));
        //设置注释内容
        comment.setString(new HSSFRichTextString("可以添加注释!"));
        //设置注释作者，当鼠标移动到单元格上是可以在状态栏中看到该内容。
        comment.setAuthor("XGT");
        //产生表格标题行
        HSSFRow row = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            HSSFCell cell = row.createCell(i);
            HSSFRichTextString text = new HSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
        //遍历集合数据，产生数据行
        Iterator<T> it = dataset.iterator();
        int index = 0;
        SimpleDateFormat sdf = new SimpleDateFormat(pattern);
        while (it.hasNext()){
            index++;
            row = sheet.createRow(index);
            T t = (T) it.next();
            for (int i = 0; i < propertys.length; i++) {
                HSSFCell cell = row.createCell(i);
                String getMethodName = "get" + propertys[i].substring(0,1).toUpperCase() + propertys[i].substring(1);
                try {
                    @SuppressWarnings("rawtypes")
                    Class tCls = t.getClass();
                    Method getMethod = tCls.getMethod(getMethodName, new Class[] {});
                    Object value = getMethod.invoke(t, new Object(){});
                    //判断值得类型后进行强制类型转换
                    String textValue = null;
                    if (value instanceof Boolean){
                        @SuppressWarnings("unused")
                        Boolean bValue =(Boolean)value;
                        textValue = value.toString();
                    }else if (value instanceof Date){
                        Date date = (Date) value;
                        textValue = sdf.format(date);
                    }else if (value instanceof byte[]){
                        //有图片时，设置行高为60px
                        row.setHeightInPoints(60);
                        //设置图片所在列宽高为80px，注意单位换算
                        sheet.setColumnWidth(i, (short)(35.7*80));
                        byte[] bsValue = (byte[]) value;
                        HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 1023, 255, (short) 6, index, (short) 6, index);
                        anchor.setAnchorType(2);
                        patriarch.createPicture(anchor,workbook.addPicture(bsValue, HSSFWorkbook.PICTURE_TYPE_JPEG));
                    }else {
                        //其他数据类型都当作字符串简单处理
                        textValue = value == null ? "" : value.toString();
                    }
                    //如果不是图片数据，就利用正则表达式判断textValue是否全部由数字组成
                    if (textValue != null){
                        Pattern p = Pattern.compile("^//d+(//.//d+)?$");
                        Matcher matcher = p.matcher(textValue);
                        if (matcher.matches()){
                            //是数字当作double处理
                            cell.setCellValue(Double.parseDouble(textValue));
                        }else {
                            HSSFRichTextString richString = new HSSFRichTextString(textValue);
                            cell.setCellValue(richString);
                        }
                    }
                } catch (NoSuchMethodException e) {
                    e.printStackTrace();
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                } catch (InvocationTargetException e) {
                    e.printStackTrace();
                }finally {
                    //清空资源
                }
            }
        }
        try{
            workbook.write(out);
            out.flush();
            out.close();
        } catch (IOException e) {
            logger.info("{}", e);
        }
    }

    public static void exportExcel2(String title, String[] headers, String[] propertys, List<Map<String,Object>> dataset,OutputStream out,String pattern){
        //声明一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //生成一个表格
        HSSFSheet sheet = workbook.createSheet(title);
        //设置表格默认列宽为个字节
        sheet.setDefaultColumnWidth((short)15);
        //产生表格标题行
        HSSFRow row = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            HSSFCell cell = row.createCell(i);
            HSSFRichTextString text = new HSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
        //遍历集合数据，产生数据行
        int rowIndex = 1;
        for (Map<String,Object> dataMap : dataset){
            row = sheet.createRow(rowIndex);
            int colIndex = 0;
            for (String property : propertys) {
                Object cd = dataMap.get(property);
                //bool值处理
                if (cd != null && cd instanceof Boolean){
                    row.createCell(colIndex, XSSFCell.CELL_TYPE_BOOLEAN).setCellValue((Boolean) cd);
                }else if (cd != null && cd instanceof Date){
                    XSSFRichTextString cv = new XSSFRichTextString(new DateTime(cd).toString(pattern));
                    row.createCell(colIndex, XSSFCell.CELL_TYPE_STRING).setCellValue(cv);
                }else {
                    //其他类型作为字符串处理
                    row.createCell(colIndex, XSSFCell.CELL_TYPE_STRING).setCellValue(cd == null ? StringUtils.EMPTY : cd.toString());
                }
                colIndex ++ ;
            }
            rowIndex ++ ;
        }
        try {
            workbook.write(out);
            out.flush();
            out.close();
        } catch (IOException e) {
            logger.info("{}", e);
        }
    }

}
