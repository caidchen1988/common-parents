package com.qtdbp.poi.excel.utils;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.model.StylesTable;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 基于XSSF and DOM (Event API)
 *
 * 读取excel的第一个Sheet的内容
 * @author: caidchen
 * @create: 2017-07-19 13:17
 * To change this template use File | Settings | File Templates.
 */
public class ReadExcelByDOMUtils {

    private int headCount = 0;
    // 开始解析行数
    private Integer startRow ;
    // 结束解析行数
    private Integer endRow ;
    // excel中所有的sheet名称
    private List<List<String>> list = new ArrayList<List<String>>();
    private static StylesTable stylesTable;

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

        return list;
    }

    @Test
    public void testProcess() throws Exception {

        String fileName = "D:\\tmp\\物流指数.xlsx";

        ReadExcelByDOMUtils howto = new ReadExcelByDOMUtils();
        List<List<String>> result = howto.processDOMReadSheet(fileName,2);
    }
}
