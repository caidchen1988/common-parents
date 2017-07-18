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
 * @author: caidchen
 * @create: 2017-07-18 9:21
 * To change this template use File | Settings | File Templates.
 */
public class DocTest {
	
    public static void main(String[] args) {
		
    	InputStream inStream;
    	OutputStream outStream;
		try {
			inStream = getInFileStream("D:\\tmp\\计费boss系统.odt");
			outStream = getOutFileStream("D:\\tmp\\odt.pdf");
			Converter converter = new Odt2PdfConverter(inStream, outStream, true, true);

			/* docx转pdf */
			/*inStream = getInFileStream("D:\\tmp\\计费boss系统.docx");
			outStream = getOutFileStream("D:\\tmp\\doxc.pdf");
			Converter converter = new Docx2PdfConverter(inStream, outStream, false, true);*/

			/* doc转pdf */
			/*inStream = getInFileStream("D:\\tmp\\计费boss系统1.doc");
			outStream = getOutFileStream("D:\\tmp\\doc.pdf");
			Converter converter = new Docx2PdfConverter(inStream, outStream, false, true);*/

			/* pptx转pdf */
			/*inStream = getInFileStream("D:\\tmp\\大数据架构文档.pptx");
			outStream = getOutFileStream("D:\\tmp\\pptx.pdf");
			Converter converter = new Pptx2PdfConverter(inStream, outStream, false, true);*/

			/* ppt转pdf */
			/*inStream = getInFileStream("D:\\tmp\\大数据架构文档1.ppt");
			outStream = getOutFileStream("D:\\tmp\\ppt1.pdf");
			Converter converter = new Pptx2PdfConverter(inStream, outStream, true, true);*/

			converter.convert();
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
