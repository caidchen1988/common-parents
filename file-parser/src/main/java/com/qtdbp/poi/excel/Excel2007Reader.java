package com.qtdbp.poi.excel;

import com.github.tobato.fastdfs.service.FastFileStorageClient;
import com.qtdbp.poi.excel.model.ExcelCellModel;
import com.qtdbp.poi.excel.model.ExcelRowModel;
import com.qtdbp.poi.excel.model.ExcelSheetModel;
import com.qtdbp.poi.excel.service.IExcelReader;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.springframework.beans.factory.annotation.Autowired;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.InputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * 抽象Excel2007读取器，excel2007的底层数据结构是xml文件，采用SAX的事件驱动的方法解析 xml，需要继承DefaultHandler，在遇到文件内容时，事件会触发，这种做法可以大大降低内存的耗费，特别使用于大数据量的文件。
 *
 * @author: caidchen
 * @create: 2017-05-23 8:38
 * To change this template use File | Settings | File Templates.
 */
public class Excel2007Reader  extends DefaultHandler {
    //共享字符串表
    private SharedStringsTable sst;
    //上一次的内容
    private String lastContents;
    private boolean nextIsString;

    // 存储每个sheet数据
    private List<ExcelRowModel> rowList = new ArrayList<ExcelRowModel>();
    private List<ExcelCellModel> cellList = new ArrayList<ExcelCellModel>();
    //当前行
    private int curRow = 0;
    //当前列
    private int curCol = -1;
//    private int curCol1 = 0;
    //日期标志
    private boolean dateFlag;
    //数字标志
    private boolean numberFlag;

    private boolean isTElement;

    private IExcelReader rowReader;

    public void setRowReader(IExcelReader rowReader){
        this.rowReader = rowReader;
    }

    /**
     * 用于外层调用时注入fastdfs-client上传对象
     */
    @Autowired
    private FastFileStorageClient storageClient ;

    /**只遍历一个电子表格，其中sheetId为要遍历的sheet索引，从1开始，1-3
     * @param filename
     * @param sheetId
     * @throws Exception
     */
    public void processOneSheet(String filename,int sheetId) throws Exception {
        OPCPackage pkg = OPCPackage.open(filename);
        XSSFReader r = new XSSFReader(pkg);
        SharedStringsTable sst = r.getSharedStringsTable();
        XMLReader parser = fetchSheetParser(sst);

        // 根据 rId# 或 rSheet# 查找sheet
        InputStream sheet2 = r.getSheet("rId"+sheetId);
        InputSource sheetSource = new InputSource(sheet2);
        parser.parse(sheetSource);
        sheet2.close();
    }

    /**
     * 遍历工作簿中所有的电子表格
     * @param filename 本地文件
     * @throws Exception
     */
    public void process(String filename) throws Exception {
        OPCPackage pkg = OPCPackage.open(filename);
        XSSFReader r = new XSSFReader(pkg);
        SharedStringsTable sst = r.getSharedStringsTable();
        XMLReader parser = fetchSheetParser(sst);
        XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) r.getSheetsData();
        while (sheets.hasNext()) {
            curRow = 0;
            InputStream sheet = sheets.next();
            InputSource sheetSource = new InputSource(sheet);
            parser.parse(sheetSource);

            rowReader.getSheet(new ExcelSheetModel(sheets.getSheetName(), rowList)) ;
            rowList.clear();

            sheet.close();
        }
    }

    /**
     * 遍历工作簿中所有的电子表格
     * @param stream 文件流
     * @throws Exception
     */
    public void process(InputStream stream) throws Exception {
        OPCPackage pkg = OPCPackage.open(stream);
        XSSFReader r = new XSSFReader(pkg);
        SharedStringsTable sst = r.getSharedStringsTable();
        XMLReader parser = fetchSheetParser(sst);
        XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) r.getSheetsData();
        while (sheets.hasNext()) {
            curRow = 0;
            InputStream sheet = sheets.next();
            InputSource sheetSource = new InputSource(sheet);
            parser.parse(sheetSource);

            rowReader.getSheet(new ExcelSheetModel(sheets.getSheetName(), rowList)) ;
            rowList.clear();

            sheet.close();
        }
    }

    public XMLReader fetchSheetParser(SharedStringsTable sst)
            throws SAXException {
//        XMLReader parser = XMLReaderFactory
//                .createXMLReader("org.apache.xerces.parsers.SAXParser");
        XMLReader parser = XMLReaderFactory.createXMLReader();
        this.sst = sst;
        parser.setContentHandler(this);
        return parser;
    }

    /**
     * 开始解析节点
     *
     * 继承的DefaultHandler处理类就是专门来解析xml
     * @param uri
     * @param localName
     * @param name
     * @param attributes
     * @throws SAXException
     */
    public void startElement(String uri, String localName, String name,
                             Attributes attributes) throws SAXException {


//        System.out.println("#############################");
//        System.out.println("uri: "+uri);
//        System.out.println("localName: "+localName);
//        System.out.println("name: "+name);
//        System.out.println("attributes: "+attributes);

        // c => 单元格
        if ("c".equals(name)) {
            curCol++;
            // 如果下一个元素是 SST 的索引，则将nextIsString标记为true
            String cellType = attributes.getValue("t");
            if ("s".equals(cellType)) {
                nextIsString = true;
            } else {
                nextIsString = false;
            }
            //日期格式
            String cellDateType = attributes.getValue("s");
            if ("1".equals(cellDateType)){
                dateFlag = true;
            } else {
                dateFlag = false;
            }
            String cellNumberType = attributes.getValue("s");
            if("2".equals(cellNumberType)){
                numberFlag = true;
            } else {
                numberFlag = false;
            }

        }
        //当元素为t时
        if("t".equals(name)){
            isTElement = true;
        } else {
            isTElement = false;
        }

        // 置空
        lastContents = "";
    }

    /**
     * 结束解析节点
     * 这是节点解析完成时调用的，这里我们遇到item的时候才调用下面的
     * @param uri
     * @param localName
     * @param name
     * @throws SAXException
     */
    public void endElement(String uri, String localName, String name)
            throws SAXException {

        // 根据SST的索引值的到单元格的真正要存储的字符串
        // 这时characters()方法可能会被调用多次
        if (nextIsString) {
            try {
                if(StringUtils.isNumeric(lastContents)) {
                    int idx = Integer.parseInt(lastContents);
                    lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
                }
            } catch (Exception e) {
                e.printStackTrace() ;
            }
        }

//        System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$");
//        System.out.println("endElement: " + localName + ", " + name+" , lastContents: "+ lastContents);

        // t元素也包含字符串
        if (isTElement) {
            String value = lastContents.trim();
            cellList.add(new ExcelCellModel(curCol, value));
            isTElement = false;
            // v => 单元格的值，如果单元格是字符串则v标签的值为该字符串在SST中的索引
            // 将单元格内容加入rowlist中，在这之前先去掉字符串前后的空白符
        } else if ("v".equals(name)) {
            String value = lastContents.trim();
            value = value.equals("") ? " " : value;
            try {
                // 日期格式处理
                if (dateFlag && StringUtils.isNumeric(value)) {
                    Date date = HSSFDateUtil.getJavaDate(Double.valueOf(value));
                    SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
                    value = dateFormat.format(date);
                }
                // 数字类型处理
                if (numberFlag) {
                    BigDecimal bd = new BigDecimal(value);
                    value = bd.setScale(3, BigDecimal.ROUND_UP).toString();
                }
            } catch (Exception e) {
                // 转换失败仍用读出来的值
                e.printStackTrace() ;
            }
            cellList.add(new ExcelCellModel(curCol, value));
        } else {
            // 如果标签名称为 row ，这说明已到行尾，调用 optRows() 方法
            if (name.equals("row")) {
//                rowReader.getRows(new ExcelRowModel(sheetName, curRow, cellList)) ; // sheetIndex + 1, curRow, cellList);
//                cellList.clear();
                rowList.add(new ExcelRowModel(curRow, cellList));
                cellList = new ArrayList<ExcelCellModel>();

                curRow++;
                curCol = -1;
            }
        }

    }

    /**
     * 保存节点内容
     * @param ch
     * @param start
     * @param length
     * @throws SAXException
     */
    public void characters(char[] ch, int start, int length)
            throws SAXException {
        //得到单元格内容的值
        lastContents += new String(ch, start, length);
    }
}