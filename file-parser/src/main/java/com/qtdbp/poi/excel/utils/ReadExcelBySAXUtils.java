package com.qtdbp.poi.excel.utils;

import com.qtdbp.poi.excel.model.ExcelCellModel;
import com.qtdbp.poi.excel.model.ExcelRowModel;
import com.qtdbp.poi.excel.model.ExcelSheetModel;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.junit.Test;
import org.xml.sax.*;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * 基于XSSF and SAX (Event API)
 *
 * 读取excel的第一个Sheet的内容
 * @author: caidchen
 * @create: 2017-07-19 13:17
 * To change this template use File | Settings | File Templates.
 */
public class ReadExcelBySAXUtils {

    // 开始解析行数
    private Integer startRow ;
    // 结束解析行数
    private Integer endRow ;
    private static StylesTable stylesTable;
    // 存储每个sheet数据
    private List<ExcelRowModel> rowList = new ArrayList<ExcelRowModel>();

    private XSSFReader reader;

    public void init(InputStream inputStream) {
        try {
            this.reader = reader(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (OpenXML4JException e) {
            e.printStackTrace();
        }
    }

    /**
     * 读取文件流
     * @param inputStream
     * @return
     * @throws IOException
     * @throws OpenXML4JException
     */
    private XSSFReader reader(InputStream inputStream) throws IOException, OpenXML4JException {
        OPCPackage pkg = OPCPackage.open(inputStream);
        return new XSSFReader( pkg ) ;
    }

    /**
     * 采用SAX进行解析
     * 获取excel表格中所有的sheet工作簿
     * @return {sheet_index, sheet_name}
     * @throws IOException
     * @throws OpenXML4JException
     * @throws SAXException
     */
    public List<ExcelSheetModel> processSAXReadSheet() throws IOException, OpenXML4JException, SAXException   {

        List<ExcelSheetModel> sheetContainer = new ArrayList<>();

//        XSSFReader reader = reader( inputStream );
        XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) this.reader.getSheetsData();

        int i = 0 ;
        while (sheets.hasNext()) {
            i++ ;
            sheets.next();

            sheetContainer.add(new ExcelSheetModel(sheets.getSheetName(),null,i)) ;
        }

        return sheetContainer ;
    }

    /**
     * 获取单个sheet工作簿内容
     * @param sheetId sheet index 默认从1开始
     * @param startRow 开始第几行解析，默认从1开始
     * @param endRow 结束第几行解析，null返回所有，注意：传值会导致解析变慢，数据量参超过2000时，建议传null
     * @return
     * @throws Exception
     */
    public List<ExcelRowModel> processSAXReadOneSheet(Integer sheetId,Integer startRow, Integer endRow) throws OpenXML4JException, IOException, SAXException {

        this.startRow = startRow == null ? 1 : startRow;
        this.endRow = endRow ;
//        this.list = new ArrayList<>();
        this.rowList = new ArrayList<>();

        if(sheetId == null) sheetId = 1 ;

//        XSSFReader r = reader( inputStream );
        this.stylesTable = this.reader.getStylesTable() ;

        SharedStringsTable sst = this.reader.getSharedStringsTable();
        XMLReader parser = fetchSheetParser(sst);
        InputStream sheet = this.reader.getSheet("rId"+sheetId);
        InputSource sheetSource = new InputSource(sheet);
        // 解析
        parser.parse(sheetSource);
        // 关闭
        sheet.close();

        return rowList ;
    }

    private XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
        XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
        ContentHandler handler = new SheetHandler(sst);
        parser.setContentHandler(handler);
        return parser;
    }

    /**
     * 自定义解析处理器
     * See org.xml.sax.helpers.DefaultHandler javadocs
     */
    private class SheetHandler extends DefaultHandler {

        private SharedStringsTable sst;
        private String lastContents;
        private boolean nextIsString;

        private List<ExcelCellModel> cellList = new ArrayList<ExcelCellModel>();
        private int curRow = 0;
        private int curCol = -1;

        //定义前一个元素和当前元素的位置，用来计算其中空的单元格数量，如A6和A8等
        private String preRef = null, ref = null;
        //定义该文档一行最大的单元格数，用来补全一行最后可能缺失的单元格
        private String maxRef = null;

        private CellDataType nextDataType = CellDataType.SSTINDEX;
        private final DataFormatter formatter = new DataFormatter();
        private short formatIndex;
        private String formatString;

        private SheetHandler(SharedStringsTable sst) {
            this.sst = sst;
        }

        /**
         * 解析一个element的开始时触发事件
         * @param uri
         * @param localName
         * @param name
         * @param attributes
         * @throws SAXException
         */
        public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {

            // c => cell
            if(name.equals("c")) {
                curCol++;

                //前一个单元格的位置
                if(preRef == null){
                    preRef = attributes.getValue("r");
                }else{
                    preRef = ref;
                }
                //当前单元格的位置
                ref = attributes.getValue("r");

                this.setNextDataType(attributes);

                // Figure out if the value is an index in the SST
                String cellType = attributes.getValue("t");
                if(cellType != null && cellType.equals("s")) {
                    nextIsString = true;
                } else {
                    nextIsString = false;
                }
            }

            // Clear contents cache
            lastContents = "";

        }

        /**
         * 根据element属性设置数据类型
         * @param attributes
         */
        public void setNextDataType(Attributes attributes){

            nextDataType = CellDataType.NUMBER;
            formatIndex = -1;
            formatString = null;
            String cellType = attributes.getValue("t");
            String cellStyleStr = attributes.getValue("s");
            String columData = attributes.getValue("r");

            if ("b".equals(cellType)){
                nextDataType = CellDataType.BOOL;
            }else if ("e".equals(cellType)){
                nextDataType = CellDataType.ERROR;
            }else if ("inlineStr".equals(cellType)){
                nextDataType = CellDataType.INLINESTR;
            }else if ("s".equals(cellType)){
                nextDataType = CellDataType.SSTINDEX;
            }else if ("str".equals(cellType)){
                nextDataType = CellDataType.FORMULA;
            }

            /*if (cellStyleStr != null){
                int styleIndex = Integer.parseInt(cellStyleStr);
                XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
                formatIndex = style.getDataFormat();
                formatString = style.getDataFormatString();
                if ("m/d/yy" == formatString){
                    nextDataType = CellDataType.DATE;
                    //full format is "yyyy-MM-dd hh:mm:ss.SSS";
                    formatString = "yyyy-MM-dd";
                }
                if (formatString == null){
                    nextDataType = CellDataType.NULL;
                    formatString = BuiltinFormats.getBuiltinFormat(formatIndex);
                }
            }*/

            if (cellStyleStr != null) {
                int styleIndex = Integer.parseInt(cellStyleStr);
                XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
                formatIndex = style.getDataFormat();
                formatString = style.getDataFormatString();

//                System.out.println("formatString: "+ formatString);

                if (formatString == null) {
                    nextDataType = CellDataType.NULL;
                    formatString = BuiltinFormats.getBuiltinFormat(formatIndex);
                } else if (formatString.contains("yyyy")) {
                    nextDataType = CellDataType.DATE;
                    formatString = "yyyy-MM-dd hh:mm:ss";
                }
            }
        }

        /**
         * 解析一个element元素结束时触发事件
         * @param uri
         * @param localName
         * @param name
         * @throws SAXException
         */
        public void endElement(String uri, String localName, String name) throws SAXException {
            // Process the last contents as required.
            // Do now, as characters() may be called more than once
            if(nextIsString) {
                int idx = Integer.parseInt(lastContents);
                lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
                nextIsString = false;
            }

            if (name.equals("v")) {
                // v => 单元格的值，如果单元格是字符串则v标签的值为该字符串在SST中的索引
                // 将单元格内容加入rowlist中，在这之前先去掉字符串前后的空白符

                String value = this.getDataValue(lastContents.trim(), "");
                //补全单元格之间的空单元格

               /* System.out.println("##########ref:"+ref+", preRef:"+preRef+", name:"+name);

                if(!ref.equals(preRef)){
                    int len = countNullCell(ref, preRef);
                    for(int i=0;i<len;i++){
                        cellList.add(new ExcelCellModel(curCol, null));
                    }
                }*/

                cellList.add(new ExcelCellModel(curCol, value));
            } else {
                //如果标签名称为 row，这说明已到行尾，调用 optRows() 方法
                if (name.equals("row")) {
                    //默认第一行为表头，以该行单元格数目为最大数目
                    if (curRow == 0) {
                        maxRef = ref;
                    }
                    //补全一行尾部可能缺失的单元格
                    /*if (maxRef != null) {
                        int len = countNullCell(maxRef, ref);
                        for (int i = 0; i <= len; i++) {
                            cellList.add(new ExcelCellModel(curCol, null));
                        }
                    }*/

                    rowList.add(new ExcelRowModel(curRow, cellList, curCol+1));
                    curRow++;

                    //一行的末尾重置一些数据
                    cellList = new ArrayList<>();
                    curCol = -1;
                    preRef = null;
                    ref = null;
                }
            }
        }

        /**
         * 根据数据类型获取数据
         * @param value
         * @param thisStr
         * @return
         */
        public String getDataValue(String value, String thisStr)

        {
            switch (nextDataType)
            {
                //这几个的顺序不能随便交换，交换了很可能会导致数据错误
                case BOOL:
                    char first = value.charAt(0);
                    thisStr = first == '0' ? "FALSE" : "TRUE";
                    break;
                case ERROR:
                    thisStr = "\"ERROR:" + value.toString() + '"';
                    break;
                case FORMULA:
                    thisStr = '"' + value.toString() + '"';
                    break;
                case INLINESTR:
                    XSSFRichTextString rtsi = new XSSFRichTextString(value.toString());
                    thisStr = rtsi.toString();
                    rtsi = null;
                    break;
                case SSTINDEX:
                    String sstIndex = value.toString();
                    thisStr = value.toString();
                    break;
                case NUMBER:
                    if (formatString != null){
                        thisStr = formatter.formatRawCellContents(Double.parseDouble(value), formatIndex, formatString).trim();
                    }else{
                        thisStr = value;
                    }
                    thisStr = thisStr.replace("_", "").trim();
                    break;
                case DATE:
                    try{
                        thisStr = formatter.formatRawCellContents(Double.parseDouble(value), formatIndex, formatString);
                    }catch(NumberFormatException ex){
                        thisStr = value.toString();
                    }
//                    thisStr = thisStr.replace(" ", "");
                    break;
                default: //自定义单元格格式
                    thisStr = "未知格式数据";
                    break;
            }

//            System.out.println("thisStr: "+ thisStr);
            return thisStr;
        }

        /**
         * 获取element的文本数据
         * @param ch
         * @param start
         * @param length
         * @throws SAXException
         */
        public void characters(char[] ch, int start, int length) throws SAXException {
            lastContents += new String(ch, start, length);
        }

        /**
         * 计算两个单元格之间的单元格数目(同一行)
         * @param ref
         * @param preRef
         * @return
         */
        public int countNullCell(String ref, String preRef){
            //excel2007最大行数是1048576，最大列数是16384，最后一列列名是XFD
            String xfd = ref.replaceAll("\\d+", "");
            String xfd_1 = preRef.replaceAll("\\d+", "");

            xfd = fillChar(xfd, 3, '@', true);
            xfd_1 = fillChar(xfd_1, 3, '@', true);

            char[] letter = xfd.toCharArray();
            char[] letter_1 = xfd_1.toCharArray();
            int res = (letter[0]-letter_1[0])*26*26 + (letter[1]-letter_1[1])*26 + (letter[2]-letter_1[2]);
            return res-1;
        }

        /**
         * 字符串的填充
         * @param str
         * @param len
         * @param let
         * @param isPre
         * @return
         */
        String fillChar(String str, int len, char let, boolean isPre){
            int len_1 = str.length();
            if(len_1 <len){
                if(isPre){
                    for(int i=0;i<(len-len_1);i++){
                        str = let+str;
                    }
                }else{
                    for(int i=0;i<(len-len_1);i++){
                        str = str+let;
                    }
                }
            }
            return str;
        }
    }

    @Test
    public void testProcess() throws Exception {


        String fileName = "D:\\tmp\\物流指数.xlsx";
        InputStream inputStream = new BufferedInputStream(new FileInputStream(new File(fileName))) ;

        ReadExcelBySAXUtils howto = new ReadExcelBySAXUtils();
        howto.init(inputStream);

        long start = System.currentTimeMillis() ;
        List<ExcelSheetModel> sheetContainer = howto.processSAXReadSheet() ;

        for(ExcelSheetModel sheet : sheetContainer) {
            System.out.println("工作簿:"+sheet.getSheetNum()+"，"+ sheet.getName());
        }
        long end = System.currentTimeMillis() ;
        System.out.println("###################################时间："+(end-start));


//        inputStream = new BufferedInputStream(new FileInputStream(new File(fileName))) ;
        start = System.currentTimeMillis() ;
        List<ExcelRowModel> list = howto.processSAXReadOneSheet(1,null ,4) ;
        for(ExcelRowModel row : list) {
            System.out.println("行数据：【行数："+row.getRowNum()+"，列数："+row.getTotalCellNum()+"】");
            for(ExcelCellModel cell : row.getCellList()) {
                System.out.print("单元格: 【x:"+row.getRowNum()+",y:"+cell.getColNum()+", index: "+row.getCellList().indexOf(cell)+", val:"+cell.getColVal()+"】");
            }
            System.out.println();
        }
        end = System.currentTimeMillis() ;

        System.out.println("###################################解析单个sheet时间："+(end-start));
    }
}
