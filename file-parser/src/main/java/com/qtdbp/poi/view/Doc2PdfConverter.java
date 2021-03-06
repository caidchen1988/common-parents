package com.qtdbp.poi.view;

import org.docx4j.Docx4J;
import org.docx4j.convert.in.Doc;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintStream;

/**
 * 预览word2003文档转pdf
 *
 * @author: caidchen
 * @create: 2017-07-18 9:21
 * To change this template use File | Settings | File Templates.
 */
public class Doc2PdfConverter extends Converter {

	public Doc2PdfConverter(InputStream inStream, OutputStream outStream, boolean showMessages, boolean closeStreamsWhenComplete) {
		super(inStream, outStream, showMessages, closeStreamsWhenComplete);
	}


	@Override
	public void convert() throws Exception{

		loading();

		InputStream iStream = inStream;

		WordprocessingMLPackage wordMLPackage = getMLPackage(iStream);

		processing();
		Docx4J.toPDF(wordMLPackage, outStream);

		finished();
	}

	protected WordprocessingMLPackage getMLPackage(InputStream iStream) throws Exception{
		PrintStream originalStdout = System.out;
		
		//Disable stdout temporarily as Doc convert produces alot of output
		System.setOut(new PrintStream(new OutputStream() {
			public void write(int b) {
				//DO NOTHING
			}
		}));

		WordprocessingMLPackage mlPackage = Doc.convert(iStream);
		
		System.setOut(originalStdout);
		return mlPackage;
	}

}
