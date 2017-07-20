package com.qtdbp.poi.excel.utils;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.junit.Test;
import org.xml.sax.*;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 基于XSSF and SAX (Event API)
 *
 * 读取excel的第一个Sheet的内容
 * @author: caidchen
 * @create: 2017-07-19 13:17
 * To change this template use File | Settings | File Templates.
 */
public class ReadExcelUtils {

    private int headCount = 0;
    // 开始解析行数
    private Integer startRow ;
    // 结束解析行数
    private Integer endRow ;
    // excel中所有的sheet名称
    private List<List<String>> list = new ArrayList<List<String>>();
    private static final Log log = LogFactory.getLog(ReadExcelUtils.class);

    /**
     * 通过文件流构建DOM进行解析
     * @param ins
     * @param headRowCount   跳过读取的表头的行数
     * @return
     * @throws InvalidFormatException
     * @throws IOException
     */
    public  List<List<String>> processDOMReadSheet(InputStream ins, int headRowCount) throws InvalidFormatException, IOException {
        Workbook workbook = WorkbookFactory.create(ins);
        return this.processDOMRead(workbook, headRowCount);
    }

    /**
     * 采用DOM的形式进行解析
     * @param filename
     * @param headRowCount   跳过读取的表头的行数
     * @return
     * @throws IOException
     * @throws InvalidFormatException
     * @throws Exception
     */
    public  List<List<String>> processDOMReadSheet(String filename,int headRowCount) throws InvalidFormatException, IOException {
        Workbook workbook = WorkbookFactory.create(new File(filename));
        return this.processDOMRead(workbook, headRowCount);
    }

    /**
     * 采用SAX进行解析
     * @param filename
     * @param headRowCount
     * @return
     * @throws OpenXML4JException
     * @throws IOException
     * @throws SAXException
     * @throws Exception
     */
    public List<List<String>> processSAXReadSheet(String filename,int headRowCount) throws IOException, OpenXML4JException, SAXException   {
        headCount = headRowCount;

        OPCPackage pkg = OPCPackage.open(filename);
        XSSFReader r = new XSSFReader( pkg );
        SharedStringsTable sst = r.getSharedStringsTable();
        XMLReader parser = fetchSheetParser(sst);

        Iterator<InputStream> sheets = r.getSheetsData();
        InputStream sheet = sheets.next();
        InputSource sheetSource = new InputSource(sheet);
        parser.parse(sheetSource);
        sheet.close();

//        log.debug("时间:"+ DateUtils.getNowTime()+",共读取了execl的记录数为 :"+list.size());
        log.debug("时间:"+ new Date() +",共读取了execl的记录数为 :"+list.size());

        return list;
    }

    /**
     * 采用SAX进行解析
     * 获取excel表格中所有的sheet工作簿
     * @param filename 文件
     * @return {sheet_index, sheet_name}
     * @throws IOException
     * @throws OpenXML4JException
     * @throws SAXException
     */
    public Map<Integer,String> processSAXReadSheet(String filename) throws IOException, OpenXML4JException, SAXException   {

        Map<Integer,String> sheetContainer = new HashMap();

        OPCPackage pkg = OPCPackage.open(filename);
        XSSFReader reader = new XSSFReader( pkg );
        XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) reader.getSheetsData();

        int i = 0 ;
        while (sheets.hasNext()) {
            i++ ;
            sheets.next();
            sheetContainer.put(i, sheets.getSheetName()) ;
        }

        return sheetContainer ;
    }

    /**
     * 获取单个sheet工作簿内容
     * @param filename 文件
     * @param sheetId sheet index 默认从1开始
     * @param startRow 开始第几行解析，默认从1开始
     * @param endRow 结束第几行解析，null返回所有，注意：传值会导致解析变慢，数据量参超过2000时，建议传null
     * @return
     * @throws Exception
     */
    public List<List<String>> processSAXReadOneSheet(String filename, Integer sheetId,Integer startRow, Integer endRow) throws OpenXML4JException, IOException, SAXException {

        this.startRow = startRow == null ? 1 : startRow;
        this.endRow = endRow ;
        this.list = new ArrayList<>();
        if(sheetId == null) sheetId = 1 ;

        OPCPackage pkg = OPCPackage.open(filename);
        XSSFReader r = new XSSFReader( pkg );
        SharedStringsTable sst = r.getSharedStringsTable();
        XMLReader parser = fetchSheetParser(sst);
        InputStream sheet = r.getSheet("rId"+sheetId);
        InputSource sheetSource = new InputSource(sheet);
        // 解析
        parser.parse(sheetSource);
        // 关闭
        sheet.close();

        return list ;
    }

    private XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
        XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
        ContentHandler handler = new SheetHandler(sst);
        parser.setContentHandler(handler);
        return parser;
    }

    /**
     * SAX 解析excel
     */
    private class SheetHandler extends DefaultHandler {

        private SharedStringsTable sst;
        private String lastContents;
        private boolean nextIsString;
        private boolean isNullCell;
        //读取行的索引
        private int rowIndex = 0;
        //是否重新开始了一行
        private boolean curRow = false;
        //
        private boolean flag = false;
        private List<String> rowContent;

        private SheetHandler(SharedStringsTable sst) {
            this.sst = sst;
        }

        /* 开始解析文档 */
        public void startDocument () {
            this.flag = false;
        }

        /* 开始解析节点 */
        public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
            //节点的类型
//            System.out.println("---------begin:" + rowIndex);
            if(name.equals("row")){
                rowIndex++;
            }
            //表头的行直接跳过
            this.flag = endRow == null ? rowIndex >= startRow : rowIndex >= startRow && rowIndex <= endRow ;
//            System.out.println("---------flag:" + this.flag);
            if(this.flag){
                curRow = true;
                // c => cell
                if(name.equals("c")) {
                    String cellType = attributes.getValue("t");
                    if(null == cellType){
                        isNullCell = true;
                    }else{
                        if(cellType.equals("s")) {
                            nextIsString = true;
                        } else {
                            nextIsString = false;
                        }
                        isNullCell = false;
                    }
                }
                // Clear contents cache
                lastContents = "";
            }
        }

        /* 结束解析节点 */
        public void endElement(String uri, String localName, String name)
                throws SAXException {
//            System.out.println("-------end："+rowIndex);
            if(this.flag){
                if(nextIsString) {
                    int idx = Integer.parseInt(lastContents);
                    lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
                    nextIsString = false;
                }
                if(name.equals("v")) {
                    //System.out.println(lastContents);
                    if(curRow){
                        //是新行则new一行的对象来保存一行的值
                        if(null==rowContent){
                            rowContent = new ArrayList<String>();
                        }
                        rowContent.add(lastContents);
                    }
                }else if(name.equals("c") && isNullCell){
                    if(curRow){
                        //是新行则new一行的对象来保存一行的值
                        if(null==rowContent){
                            rowContent = new ArrayList<String>();
                        }
                        rowContent.add(null);
                    }
                }

                isNullCell = false;

                if("row".equals(name)){
                    list.add(rowContent);
                    curRow = false;
                    rowContent = null;
                }
            }

        }

        /* 保存节点内容 */
        public void characters(char[] ch, int start, int length)
                throws SAXException {
            lastContents += new String(ch, start, length);
        }
    }

    /**
     * DOM的形式解析execl
     * @param workbook
     * @param headRowCount
     * @return
     * @throws InvalidFormatException
     * @throws IOException
     */
    private List<List<String>> processDOMRead(Workbook workbook,int headRowCount) throws InvalidFormatException, IOException {
        headCount = headRowCount;

        Sheet sheet = workbook.getSheetAt(0);
        //行数
        int endRowIndex = sheet.getLastRowNum();

        Row row = null;
        List<String> rowList = null;

        for(int i=headCount; i<=endRowIndex; i++){
            rowList = new ArrayList<String>();
            row = sheet.getRow(i);
            for(int j=0; j<row.getLastCellNum();j++){
                if(null==row.getCell(j)){
                    rowList.add(null);
                    continue;
                }
                int dataType = row.getCell(j).getCellType();
                if(dataType == Cell.CELL_TYPE_NUMERIC){
                    DecimalFormat df = new DecimalFormat("0.####################");
                    rowList.add(df.format(row.getCell(j).getNumericCellValue()));
                }else if(dataType == Cell.CELL_TYPE_BLANK){
                    rowList.add(null);
                }else if(dataType == Cell.CELL_TYPE_ERROR){
                    rowList.add(null);
                }else{
                    //这里的去空格根据自己的情况判断
                    String valString = row.getCell(j).getStringCellValue();
                    Pattern p = Pattern.compile("\\s*|\t|\r|\n");
                    Matcher m = p.matcher(valString);
                    valString = m.replaceAll("");
                    //去掉狗日的不知道是啥东西的空格
                    if(valString.indexOf(" ")!=-1){
                        valString = valString.substring(0, valString.indexOf(" "));
                    }

                    rowList.add(valString);
                }
            }

            list.add(rowList);
        }
//        log.debug("时间:"+DateUtils.getNowTime()+",共读取了execl的记录数为 :"+list.size());
        log.debug("时间:"+ new Date() +",共读取了execl的记录数为 :"+list.size());

        return list;
    }

    @Test
    public void testProcess() throws Exception {

        ReadExcelUtils howto = new ReadExcelUtils();
        String fileName = "D:\\tmp\\物流指数.xlsx";
        /*List<List<String>> list = howto.processSAXReadSheet(fileName,0);
        for(List<String> row : list) {
            for(String cell : row) {
                System.out.print("单元格:"+ cell +" ");
            }
            System.out.println();
        }*/

        long start = System.currentTimeMillis() ;
        Map<Integer,String> sheetContainer = howto.processSAXReadSheet(fileName) ;

        for(Map.Entry<Integer,String> sheet : sheetContainer.entrySet()) {
            System.out.println("工作簿:"+sheet.getKey()+"，"+ sheet.getValue());

        }
        long end = System.currentTimeMillis() ;
        System.out.println("###################################时间："+(end-start));

        start = System.currentTimeMillis() ;
        List<List<String>> list = howto.processSAXReadOneSheet(fileName ,2,null ,4) ;
        for(List<String> row : list) {
            for(String cell : row) {
                System.out.print("单元格:"+ cell +" ");
            }
            System.out.println();
        }
        end = System.currentTimeMillis() ;

        System.out.println("###################################解析单个sheet时间："+(end-start));
        /*ReadExcelUtils h = new ReadExcelUtils();
        String fileName1 = "f:/test.xls";
        List<List<String>> result = h.processDOMReadSheet(fileName1,2);*/
    }
}
