package com.qtdbp.poi.zip;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.tools.zip.ZipEntry;
import org.apache.tools.zip.ZipFile;
import org.apache.tools.zip.ZipOutputStream;

import java.io.*;
import java.util.Enumeration;
import java.util.regex.Pattern;

/**
 * zip压缩包处理
 *
 * 基于ant.jar， 只支持本地文件处理
 * @author: caidchen
 * @create: 2017-06-03 10:35
 * To change this template use File | Settings | File Templates.
 */
public class ZipAntUtil {

    private Log log = LogFactory.getLog(this.getClass().getName());
    private static int BUF_SIZE = 1024;
    private static String ZIP_ENCODEING = "GBK";

    public ZipAntUtil() {
        this(1024 * 10);
    }

    public ZipAntUtil(int bufSize) {
        BUF_SIZE = bufSize;
    }

    /**
     * 压缩文件或文件夹
     *
     * @param zipFileName
     * @param inputFile
     * @throws Exception
     */
    public void zip(String zipFileName, String inputFile) throws Exception {
        zip(zipFileName, new File(inputFile));
    }

    /**
     * 压缩文件或文件夹
     *
     * @param zipFileName
     * @param inputFile
     * @throws Exception
     */
    public void zip(String zipFileName, File inputFile) throws Exception {
        // 未指定压缩文件名，默认为"ZipFile"
        if (zipFileName == null || zipFileName.equals(""))
            zipFileName = "ZipFile";

        // 添加".zip"后缀
        if (!zipFileName.endsWith(".zip"))
            zipFileName += ".zip";

        // 创建文件夹
        String path = Pattern.compile("[\\/]").matcher(zipFileName).replaceAll(File.separator);
        int endIndex = path.lastIndexOf(File.separator);
        path = path.substring(0, endIndex);
        File f = new File(path);
        f.mkdirs();
        // 开始压缩
        {
            ZipOutputStream zos = new ZipOutputStream(new BufferedOutputStream(new FileOutputStream(zipFileName)));
            zos.setEncoding(ZIP_ENCODEING);
            compress(zos, inputFile, "");
            log.debug("zip done");
            zos.close();
        }
    }

    /**
     * 解压缩zip压缩文件到指定目录
     *
     * @param unZipFileName
     * @param outputDirectory
     * @throws Exception
     */
    public void unZip(String unZipFileName, String outputDirectory) throws Exception {
        // 创建输出文件夹对象
        File outDirFile = new File(outputDirectory);
        outDirFile.mkdirs();
        // 打开压缩文件文件夹
        ZipFile zipFile = new ZipFile(unZipFileName, ZIP_ENCODEING);
        for (Enumeration entries = zipFile.getEntries(); entries.hasMoreElements();) {
            ZipEntry ze = (ZipEntry) entries.nextElement();
            File file = new File(outDirFile, ze.getName());
            if (ze.isDirectory()) {// 是目录，则创建之
                file.mkdirs();
                log.debug("mkdir " + file.getAbsolutePath());
            } else {
                File parent = file.getParentFile();
                if (parent != null && !parent.exists()) {
                    parent.mkdirs();
                }
                log.debug("unziping " + ze.getName());
                file.createNewFile();
                FileOutputStream fos = new FileOutputStream(file);
                InputStream is = zipFile.getInputStream(ze);
                this.inStream2outStream(is, fos);
                fos.close();
                is.close();
            }
        }
        zipFile.close();
    }

    /**
     * 压缩一个文件夹或文件对象到已经打开的zip输出流 <b>不建议直接调用该方法</b>
     *
     * @param zos
     * @param f
     * @param fileName
     * @throws Exception
     */
    public void compress(ZipOutputStream zos, File f, String fileName) throws Exception {
        log.debug("Zipping " + f.getName());
        if (f.isDirectory()) {
            // 压缩文件夹
            File[] fl = f.listFiles();
            zos.putNextEntry(new ZipEntry(fileName + "/"));
            fileName = fileName.length() == 0 ? "" : fileName + "/";
            for (int i = 0; i < fl.length; i++) {
                compress(zos, fl[i], fileName + fl[i].getName());
            }
        } else {
            // 压缩文件
            zos.putNextEntry(new ZipEntry(fileName));
            FileInputStream fis = new FileInputStream(f);
            this.inStream2outStream(fis, zos);
            fis.close();
            zos.closeEntry();
        }
    }

    private void inStream2outStream(InputStream is, OutputStream os) throws IOException {
        BufferedInputStream bis = new BufferedInputStream(is);
        BufferedOutputStream bos = new BufferedOutputStream(os);
        int bytesRead = 0;
        for (byte[] buffer = new byte[BUF_SIZE]; ((bytesRead = bis.read(buffer, 0, BUF_SIZE)) != -1);) {
            bos.write(buffer, 0, bytesRead); // 将流写入
        }
    }

    public static void main(String[] args) throws Exception {
        try {
            ZipAntUtil zipManger = new ZipAntUtil() ;
            zipManger.unZip("D:\\test.zip", "D:\\tmp");
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
}
