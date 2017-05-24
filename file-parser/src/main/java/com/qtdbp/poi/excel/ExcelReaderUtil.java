package com.qtdbp.poi.excel;

import com.qtdbp.poi.excel.service.IExcelReader;

import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;

/**
 * @author: caidchen
 * @create: 2017-05-23 8:40
 * To change this template use File | Settings | File Templates.
 */
public class ExcelReaderUtil {

    //excel2003扩展名
    public static final String EXCEL03_EXTENSION = ".xls";
    //excel2007扩展名
    public static final String EXCEL07_EXTENSION = ".xlsx";

    /**
     * 读取Excel文件，可能是03也可能是07版本
     * @param excel03
     * @param excel07
     * @param localFileName 本地文件路劲
     * @throws Exception
     */
    public static void readExcel(IExcelReader reader, String localFileName) throws Exception{
        // 处理excel2003文件
        if (localFileName.endsWith(EXCEL03_EXTENSION)){
            Excel2003Reader excel03 = new Excel2003Reader();
            excel03.setRowReader(reader);
            excel03.process(localFileName);
            // 处理excel2007文件
        } else if (localFileName.endsWith(EXCEL07_EXTENSION)){
            Excel2007Reader excel07 = new Excel2007Reader();
            excel07.setRowReader(reader);
            excel07.process(localFileName);
        } else {
            // throw new  Exception("文件格式错误，fileName的扩展名只能是xls或xlsx。");
        }
    }

    /**
     * 读取Excel文件，可能是03也可能是07版本
     * @param reader
     * @param fileName 文件名称
     * @param stream 文件流
     * @throws Exception
     */
    public static void readExcel(IExcelReader reader, String fileName, InputStream stream) throws Exception{
        // 处理excel2003文件
        if (fileName.endsWith(EXCEL03_EXTENSION)){
            Excel2003Reader excel03 = new Excel2003Reader();
            excel03.setRowReader(reader);
            excel03.process(stream);
            // 处理excel2007文件
        } else if (fileName.endsWith(EXCEL07_EXTENSION)){
            Excel2007Reader excel07 = new Excel2007Reader();
            excel07.setRowReader(reader);
            excel07.process(stream);
        } else {
            // throw new  Exception("文件格式错误，fileName的扩展名只能是xls或xlsx。");
        }
    }

    /**
     * 读取Excel文件，可能是03也可能是07版本
     * @param reader
     * @param remoteFileName 远程文件地址
     * @throws Exception
     */
    public static void readRemoteExcel(IExcelReader reader, String remoteFileName) throws Exception{
        // 处理excel2003文件
        if (remoteFileName.endsWith(EXCEL03_EXTENSION)){
            Excel2003Reader excel03 = new Excel2003Reader();
            excel03.setRowReader(reader);
            excel03.process(path2Stream(remoteFileName));
            // 处理excel2007文件
        } else if (remoteFileName.endsWith(EXCEL07_EXTENSION)){
            Excel2007Reader excel07 = new Excel2007Reader();
            excel07.setRowReader(reader);
            excel07.process(path2Stream(remoteFileName));
        } else {
            // throw new  Exception("文件格式错误，fileName的扩展名只能是xls或xlsx。");
        }
    }

    /**
     * 根据地址获得数据的字节流
     *
     * @param filePath
     *            网络连接地址
     * @return
     */
    public static InputStream path2Stream(String filePath) {
        try {
            URL url = new URL(filePath);
            HttpURLConnection conn = (HttpURLConnection) url.openConnection();
            conn.setRequestMethod("GET");
            conn.setConnectTimeout(5 * 1000);
            return conn.getInputStream();// 通过输入流获取文件数据
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }
}
