package com.qtdbp.poi.view;

import com.qtdbp.poi.excel.model.ExcelCellModel;
import com.qtdbp.poi.excel.model.ExcelRowModel;
import com.qtdbp.poi.excel.model.ExcelSheetModel;
import com.qtdbp.poi.excel.utils.ReadExcelBySAXUtils;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import org.apache.commons.collections.map.HashedMap;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.xml.sax.SAXException;

import java.io.*;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import static java.util.stream.Collectors.toMap;

/**
 * Excel转html
 *
 * @author: caidchen
 * @create: 2017-07-23 17:18
 * To change this template use File | Settings | File Templates.
 */
public class Xslx2HtmlConverter extends Converter {

    public Xslx2HtmlConverter(InputStream inStream, OutputStream outStream, boolean showMessages, boolean closeStreamsWhenComplete) {
        super(inStream, outStream, showMessages, closeStreamsWhenComplete);
    }

    @Override
    public void convert() throws Exception {

        loading();

        ReadExcelBySAXUtils howto = new ReadExcelBySAXUtils();
        howto.init(inStream);
        long start = System.currentTimeMillis() ;

        // 获取数据解析数据
        Map<String,Object> map = formatData(howto, page) ;

        long end = System.currentTimeMillis() ;
        System.out.println("###################################解析单个sheet时间："+(end-start));

        // 输出html
        toHtml(map) ;

        finished();
    }

    private Map<String, Object> formatData(ReadExcelBySAXUtils howto, Integer sheetId) throws OpenXML4JException, SAXException, IOException {

        Map<String,Object> map = new HashedMap() ;

        List<ExcelSheetModel> sheetList = howto.processSAXReadSheet() ;

        List<ExcelRowModel> rowList = null ;
        List<ExcelCellModel> titleCellList = null ; // 第一行标题单元格列表

        if(sheetList != null && !sheetList.isEmpty()) {
            // 默认解析第一个sheet
            if(sheetId == null) sheetId = 1 ;
            rowList = howto.processSAXReadOneSheet(sheetId, null, null);

            if(rowList != null && !rowList.isEmpty()) {

                // 处理行数据，补充不存在的列数据
                formatRowData(rowList) ;

                titleCellList = rowList.get(0).getCellList() ;

                // 删除第一行
                rowList.remove(0) ;
            }
        }

        // 解析完数据放到map中
        map.put("sheetList", sheetList) ;
        map.put("rowList", rowList) ;
        map.put("titleCellList", titleCellList) ;

        return map ;
    }

    private void formatRowData(List<ExcelRowModel> rowList) {

        for (ExcelRowModel row : rowList) {

            List<ExcelCellModel> cellList = row.getCellList() ;

            // Java 8, Convert List to Map
            Map<Integer, ExcelCellModel> cellMap = cellList.stream().collect(toMap(ExcelCellModel::getColNum, (p) -> p));

            // 补全单元格数据
            for (int i = 0; i < row.getTotalCellNum(); i++) {
                if(!cellMap.containsKey(i)) {
                    cellMap.put(i, new ExcelCellModel(i, "") ) ;
                }
            }

            // Java 8, Convert all Map values  to a List
            cellList = cellMap.values().stream().collect(Collectors.toList()) ;
            row.setCellList(cellList) ;
//            cellMap.forEach((k,v)-> cellList.add(k, v) );
        }

    }

    /**
     * freemarker输出成html
     * @param map
     */
    private void toHtml(Map<String,Object> map) {

        Template template = null ;
        try {
            //指定版本号
            Configuration configuration = new Configuration(Configuration.VERSION_2_3_0);
            //设置模板目录
            configuration.setClassForTemplateLoading(Xslx2HtmlConverter.class,"\\templates");
            //设置默认编码格式
            configuration.setDefaultEncoding("UTF-8");
            //从设置的目录中获得模板
            template = configuration.getTemplate("xlsx2html.ftl");
        } catch (IOException e) {
            e.printStackTrace();
        }

        //合并模板和数据模型
        Writer out = null ;
        try {
            out = new OutputStreamWriter(outStream);
            template.process(map, out);

            out.flush();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (TemplateException e) {
            e.printStackTrace();
        } finally {
            //关闭
            if(out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }

        }

    }
}
