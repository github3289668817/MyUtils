package per.xgt.ReponseUtil;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.common.IOUtil;
import per.xgt.IoUtil.IoUtil;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.net.URLEncoder;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * @Author: Valen
 * @Email：ValentinoXiong@163.com
 * @CreateTime: 2020-12-11 17:57:02
 * @Descirption:
 */
public class ResponseUtil {

    //response文件下载设置
    public static void ResponseSetForFile(HttpServletResponse response,String fileName){
        //文件名编码设置，防止中文乱码
        try {
            fileName = URLEncoder.encode(fileName,"UTF-8");
        } catch (Exception e){
            //抛出异常
            System.out.println(e.getMessage());
        }
        //设置文件类型，编码和文件名
        response.setHeader("content-disposition", "attachment;filename="+fileName);
        response.setCharacterEncoding("utf-8");
        response.setContentType("application/octet-stream");
    }

    //workbook输出到浏览器
    public static void writeBrowser(XSSFWorkbook workbook,HttpServletResponse response){
        try {
            //workbook输出到浏览器
            workbook.write(response.getOutputStream());
        } catch (Exception e){
            //抛出异常
            System.out.println(e.getMessage());
        }
    }

    private void exportZip(List<Workbook> workbooks,String fileName,HttpServletResponse response){
        try (ZipOutputStream zipOut = new ZipOutputStream(response.getOutputStream())){
            //excel文件循环添加到zipOut中
            ZipEntry entry;
            ByteArrayOutputStream byteOut;
            ByteArrayInputStream byteIn;
            int index =1;
            for (Workbook workbook : workbooks){
                //设置文件名字
                entry = new ZipEntry(fileName+"_"+(index++)+".xlsx");
                zipOut.putNextEntry(entry);
                //workbook写入到byteOut（内润流中），以你为workBook.write（zipOut）会将ZipOut关闭，下次调用时报ZipOut已关闭
                byteOut = new ByteArrayOutputStream();
                workbook.write(byteOut);
                byteIn = new ByteArrayInputStream(byteOut.toByteArray());
                //将文件写入到zipOut中
                IoUtil.copy(byteIn, byteOut);
            }
            //刷新zipOut缓存
            zipOut.flush();
        } catch (Exception e){
            System.out.println(e.getMessage());
        }
    }

}
