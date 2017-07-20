package com.qtdbp.poi.view;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * 预览文档测试
 *
 * 问题：如果文件强转格式，会导致解析失败
 * @author: caidchen
 * @create: 2017-07-18 9:21
 * To change this template use File | Settings | File Templates.
 */
public class DocTest {
	
    public static void main(String[] args) {

//    	String infileName = "D:\\tmp\\计费boss系统.odt" ;
//		String infileName = "D:\\tmp\\计费boss系统.docx" ;
		String infileName = "D:\\tmp\\2.doc" ;
//		String infileName = "D:\\tmp\\大数据架构文档.pptx" ;
//		String infileName = "D:\\tmp\\1.ppt" ;

		try {
			String outFileName = "D:\\tmp\\out.pdf" ;
			InputStream inStream = getInFileStream(infileName) ;
			OutputStream outStream = getOutFileStream(outFileName) ;
			ConverterUtil.process(inStream, infileName, outStream);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
    
    protected static InputStream getInFileStream(String inputFilePath) throws FileNotFoundException{
		File inFile = new File(inputFilePath);
		FileInputStream iStream = new FileInputStream(inFile);
		return iStream;
	}

	protected static OutputStream getOutFileStream(String outputFilePath) throws IOException{
		File outFile = new File(outputFilePath);
		
		try{
			//Make all directories up to specified
			outFile.getParentFile().mkdirs();
		} catch (NullPointerException e){
			//Ignore error since it means not parent directories
		}
		
		outFile.createNewFile();
		FileOutputStream oStream = new FileOutputStream(outFile);
		return oStream;
	}
}
