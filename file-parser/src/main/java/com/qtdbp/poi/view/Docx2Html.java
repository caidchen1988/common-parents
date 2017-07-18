package com.qtdbp.poi.view;

import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.core.FileURIResolver;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Test;

import java.io.*;

/**
 * 预览word2007文档转html
 *
 * @author: caidchen
 * @create: 2017-07-18 9:21
 * To change this template use File | Settings | File Templates.
 */
public class Docx2Html
{

	@Test
	 public void doGenerateSysOut( )
     throws IOException
    {
		   String fileInName = "test" ;
		   String root = "target";
	        String fileOutName =  "F://tests123123" + fileInName + ".html";
	     
	        long startTime = System.currentTimeMillis();

	        XWPFDocument document = new XWPFDocument( new FileInputStream("F:/紧急任务申请单.docx") );
	        XHTMLOptions options = XHTMLOptions.create().indent( 4 );
	        // Extract image
	        File imageFolder = new File( "D:/vfsroot/1000000/ueditor_upload/images" + fileInName );
	        options.setExtractor( new FileImageExtractor( imageFolder ) );
	        // URI resolver
	        options.URIResolver( new FileURIResolver( imageFolder ) );
	        
	        File outFile = new File( fileOutName );
	        outFile.getParentFile().mkdirs();
	        OutputStream out = new FileOutputStream( outFile );
	        XHTMLConverter.getInstance().convert( document, out, options );

	        System.out.println( "Generate " + fileOutName + " with " + ( System.currentTimeMillis() - startTime ) + " ms." );
 }


	/*public static String conver2Html(InputStream in){
    	String content = "";
    	try
        {
            XWPFDocument document = new XWPFDocument(in);
         //   String outFilePath = VFSUtil.getVFSPath( outputFileDir +File.separator+ "out.htm" );
            File outFile = new File( outFilePath );
            outFile.getParentFile().mkdirs();

            XHTMLOptions options = XHTMLOptions.create();
            
            //resolver path
            String resolverPath = VFSUtil.getVFSPath( outputFileDir );
           
            File imageFolder = new File(resolverPath);
            // File resolver
            options.setExtractor( new FileImageExtractor( imageFolder ) );
            // URI  resolver
            options.URIResolver( new FileURIResolver( imageFolder ) );
            
            OutputStream out = new FileOutputStream( outFile );
            XHTMLConverter.getInstance().convert( document, out, options );
            
            // Html
            String htmlContent = stream2String(new FileInputStream(outFile),"UTF-8");
            
            // DealWith images 
            File imageFile = new File( resolverPath +File.separator + "/word/media" );
            
            // First Rename imageFile
            Map<String,String> map = new HashMap<String, String>();
            content =  getImgStr(htmlContent,map);
            
            FileUtil.deleteFile( resolverPath+File.separator+"/word" );
            
        }
        catch ( Throwable e )
        {
            e.printStackTrace();
        }
    	return content;
    }
    
    *//**
     * File2String
     * @param in
     * @param charset
     * @return String
     *//*
    public static String stream2String(InputStream in, String charset) {
        StringBuffer sb = new StringBuffer(); 
        try { 
                Reader r = new InputStreamReader(in, charset); 
                int length = 0; 
                for (char[] c = new char[1024]; (length = r.read(c)) != -1;) { 
                        sb.append(c, 0, length); 
                } 
                r.close(); 
        } catch (UnsupportedEncodingException e) { 
                e.printStackTrace(); 
        } catch (FileNotFoundException e) { 
                e.printStackTrace(); 
        } catch (IOException e) { 
                e.printStackTrace(); 
        } 
        return sb.toString(); 
   }
    

    *//**
     * replace image src
     * @param htmlStr
     * @return String
     *//*
    public static String getImgStr(String htmlStr,Map<String,String> map){  
		 String img="";     
	     Pattern p_image;
	     Matcher m_image;
	     String content = htmlStr;
	     String regEx_img = "<img.*src=(.*?)[^>]*?>"; //图片链接地址     
	     p_image = Pattern.compile(regEx_img,Pattern.CASE_INSENSITIVE);     
	     m_image = p_image.matcher(htmlStr);   
	     while(m_image.find()){     
	         img =  m_image.group();     
	         Matcher m  = Pattern.compile("src=\"?(.*?)(\"|>|\\s+)").matcher(img); //匹配src  
	         while(m.find()){  
	        	String src = FileUtil.normalize(m.group(1));
	        	File file = new File(src);
	            if(file.exists() && file.isFile()){
	            	String orginalFileName = file.getName();
	        		String uuid = UUID.randomUUID().toString().replaceAll("-", "");
	                String fileType = FileUtil.getFileType(orginalFileName);
	                String newFileName = uuid + "." + fileType;
	                String desPath = FileUtil.normalize(outputFileDir +File.separator+ newFileName) ;
	                file.renameTo(new File( VFSUtil.getVFSPath(desPath) ));
            		content = content.replace(m.group(1), "ckfile.do?path="+desPath);
	            }
	         }  
	     } 
	     return content;     
	 } */
}