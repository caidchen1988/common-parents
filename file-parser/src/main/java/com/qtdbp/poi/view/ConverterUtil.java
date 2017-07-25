package com.qtdbp.poi.view;

import java.io.*;

/**
 * 预览转换
 *
 * 问题poi必须在3.16版本才能用ppt生成图片，但是docx必须在3.14版本才能用
 * @author: caidchen
 * @create: 2017-07-18 9:21
 * To change this template use File | Settings | File Templates.
 */
public class ConverterUtil {

	private static boolean shouldShowMessages = false;

	/**
	 * 文件转换，其他文件类型转pdf
	 *
	 * @param inStream
	 * @param inPath
	 * @param page sheet表位置，从1开始
	 * @return
	 */
	public static void process(InputStream inStream, String inPath, OutputStream outStream, Integer page) {

		try {

			String lowerCaseInPath = inPath.toLowerCase();

			Converter converter ;
			if (lowerCaseInPath.endsWith("doc")) {
				converter = new Doc2PdfConverter(inStream, outStream, shouldShowMessages, true);
			} else if (lowerCaseInPath.endsWith("docx")) {
				converter = new Docx2PdfConverter(inStream, outStream, shouldShowMessages, true);
			} else if (lowerCaseInPath.endsWith("ppt")) {
				converter = new Ppt2PdfConverter(inStream, outStream, shouldShowMessages, true);
			} else if (lowerCaseInPath.endsWith("pptx")) {
				converter = new Pptx2PdfConverter(inStream, outStream, shouldShowMessages, true);
			} else if (lowerCaseInPath.endsWith("odt")) {
				converter = new Odt2PdfConverter(inStream, outStream, shouldShowMessages, true);
			} else if (lowerCaseInPath.endsWith("xlsx")){
				converter = new Xslx2HtmlConverter(inStream, outStream, shouldShowMessages, true);
				converter.setPage(page);
			} else {
				converter = null;
			}

			if (converter != null) converter.convert();
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
}
