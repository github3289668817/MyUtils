package per.xgt.IoUtil;

import java.io.InputStream;
import java.io.OutputStream;

/**
 * @Author: Valen
 * @Email：ValentinoXiong@163.com
 * @CreateTime: 2020-12-14 10:29:15
 * @Descirption:
 */
public class IoUtil {

    public static long copy(InputStream in, OutputStream out) throws Exception{
        return copy((InputStream)in,(OutputStream)out,1024);
    }

    public static long copy(InputStream in,OutputStream out,int bufferSize) throws Exception {
        return copy((InputStream) in,(OutputStream) out,bufferSize,0);
    }

    public static long copy(InputStream in,OutputStream out,int bufferSize,int flag){
        if (bufferSize <= 0){
            bufferSize = 1024;
        }
        byte[] buffer = new byte[bufferSize];
        long size = 0L;
        try {
            boolean var7 = true;
            int readSize;
            while ((readSize = in.read(buffer)) != -1){
                out.write(buffer,0,readSize);
                size += (long)readSize;
                out.flush();
            }
        } catch (Exception e){
            System.out.println(e.getMessage());
        }
        return size;
    }

}
