package com.qtdbp.poi.view;

import org.apache.poi.hslf.usermodel.*;

import java.awt.*;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * 预览ppt2003文档转pdf
 *
 * @author: caidchen
 * @create: 2017-07-18 9:21
 * To change this template use File | Settings | File Templates.
 */
public class Ppt2PdfConverter extends Pptx2PdfConverter {

	//构造一页ppt
	private List<HSLFSlide> slides;
	
	public Ppt2PdfConverter(InputStream inStream, OutputStream outStream, boolean showMessages, boolean closeStreamsWhenComplete) {
		super(inStream, outStream, showMessages, closeStreamsWhenComplete);
	}

	protected Dimension processSlides() throws IOException{
		HSLFSlideShow ppt = new HSLFSlideShow(inStream);
		Dimension dimension = ppt.getPageSize();
		slides = ppt.getSlides();
		return dimension;
	}

	protected int getNumSlides(){
		return slides.size();
	}

	protected void drawOntoThisGraphic(int index, Graphics2D graphics){
		slides.get(index).draw(graphics);
	}

	protected Color getSlideBGColor(int index){
		return slides.get(index).getBackground().getFill().getForegroundColor();
	}

	/**
	 * 解决中文乱码
	 */
	protected void solveChineseDisorderlyCode(int index) {

		List<HSLFShape> shapes = slides.get(index).getShapes() ;
		if(shapes == null || shapes.isEmpty()) return;
		//防止中文乱码
		for( HSLFShape shape : shapes ){
			if ( shape instanceof HSLFTextShape){
				HSLFTextShape txtshape = (HSLFTextShape)shape ;

				for ( HSLFTextParagraph textPara : txtshape.getTextParagraphs() ){
					List<HSLFTextRun> textRunList = textPara.getTextRuns();
					for(HSLFTextRun textRun: textRunList) {
						textRun.setFontFamily("宋体");
					}
				}
			}
		}
	}
}
