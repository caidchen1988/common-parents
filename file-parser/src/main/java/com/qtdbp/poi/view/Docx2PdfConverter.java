package com.qtdbp.poi.view;

import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.InputStream;
import java.io.OutputStream;

/**
 * 预览word2007文档转pdf
 *
 * @author: caidchen
 * @create: 2017-07-18 9:21
 * To change this template use File | Settings | File Templates.
 */
public class Docx2PdfConverter extends Converter {

	public Docx2PdfConverter(InputStream inStream, OutputStream outStream, boolean showMessages, boolean closeStreamsWhenComplete) {
		super(inStream, outStream, showMessages, closeStreamsWhenComplete);
	}

	@Override
	public void convert() throws Exception {
		loading();       
		

        XWPFDocument document = new XWPFDocument(inStream);

        PdfOptions options = PdfOptions.create();

        
        processing();
        PdfConverter.getInstance().convert(document, outStream, options);
        
        finished();
        
	}

}
