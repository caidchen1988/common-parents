package com.qtdbp.poi.view;

import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;

import java.awt.*;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * 预览ppt2003文档转pdf
 *
 * @author: caidchen
 * @create: 2017-07-18 9:21
 * To change this template use File | Settings | File Templates.
 */
public class Ppt2PdfConverter extends Pptx2PdfConverter {

	//构造一页ppt
	private java.util.List<HSLFSlide> slides;
	
	public Ppt2PdfConverter(InputStream inStream, OutputStream outStream, boolean showMessages, boolean closeStreamsWhenComplete) {
		super(inStream, outStream, showMessages, closeStreamsWhenComplete);
	}


	@Override	
	protected Dimension processSlides() throws IOException{

		//构造ppt对象
		HSLFSlideShow ppt = new HSLFSlideShow(inStream);
		Dimension dimension = ppt.getPageSize();
		slides = ppt.getSlides();
		return dimension;
	}
	
	@Override
	protected int getNumSlides(){
		return slides.size();
	}
	
	@Override
	protected void drawOntoThisGraphic(int index, Graphics2D graphics){
		slides.get(index).draw(graphics);
	}
	
	@Override
	protected Color getSlideBGColor(int index){
		return slides.get(index).getBackground().getFill().getForegroundColor();
	}

}
