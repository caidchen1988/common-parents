package com.qtdbp.poi.view;

import com.itextpdf.text.Document;
import com.itextpdf.text.Image;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.awt.geom.AffineTransform;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * 预览ppt2007文档转pdf
 *
 * @author: caidchen
 * @create: 2017-07-18 9:21
 * To change this template use File | Settings | File Templates.
 */
public class Pptx2PdfConverter extends Converter{

	public Pptx2PdfConverter(InputStream inStream, OutputStream outStream, boolean showMessages, boolean closeStreamsWhenComplete) {
		super(inStream, outStream, showMessages, closeStreamsWhenComplete);
	}

	private List<XSLFSlide> slides;

	@Override
	public void convert() throws Exception {
		loading();
		
		Dimension pgsize = processSlides();
		
		processing();
		
	    double zoom = 2; // magnify it by 2 as typical slides are low res
	    AffineTransform at = new AffineTransform();
	    at.setToScale(zoom, zoom);

		
		Document document = new Document();

		PdfWriter writer = PdfWriter.getInstance(document, outStream);
		document.open();
		
		for (int i = 0; i < getNumSlides(); i++) {

			// 解决中文乱码问题
			solveChineseDisorderlyCode(i) ;

			BufferedImage bufImg = new BufferedImage((int)Math.ceil(pgsize.width*zoom), (int)Math.ceil(pgsize.height*zoom), BufferedImage.TYPE_INT_RGB);
			Graphics2D graphics = bufImg.createGraphics();
			graphics.setTransform(at);
			//clear the drawing area
			graphics.setPaint(getSlideBGColor(i));
			graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));
			try{
				drawOntoThisGraphic(i, graphics);
			} catch(Exception e){
				//Just ignore, draw what I have
				e.printStackTrace();
				System.out.println("第"+i+"张ppt转换出错");
			}
			
			Image image = Image.getInstance(bufImg, null);
			document.setPageSize(new Rectangle(image.getScaledWidth(), image.getScaledHeight()));
			document.newPage();
			image.setAbsolutePosition(0, 0);
			document.add(image);
		}
		//Seems like I must close document if not output stream is not complete
		document.close();
		
		//Not sure what repercussions are there for closing a writer but just do it.
		writer.close();
		finished();
		

	}
	
	protected Dimension processSlides() throws IOException{
		InputStream iStream = inStream;
		XMLSlideShow ppt = new XMLSlideShow(iStream);
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
		return slides.get(index).getBackground().getFillColor();
	}

	/**
	 * 解决中文乱码
	 */
	private void solveChineseDisorderlyCode(int index) {

		List<XSLFShape> shapes = slides.get(index).getShapes() ;
		if(shapes == null || shapes.isEmpty()) return;
		//防止中文乱码
		for( XSLFShape shape : shapes ){
			if ( shape instanceof XSLFTextShape){
				XSLFTextShape txtshape = (XSLFTextShape)shape ;

				for ( XSLFTextParagraph textPara : txtshape.getTextParagraphs() ){
					List<XSLFTextRun> textRunList = textPara.getTextRuns();
					for(XSLFTextRun textRun: textRunList) {
						textRun.setFontFamily("宋体");
					}
				}
			}
		}
	}

}
