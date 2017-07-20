package com.qtdbp.poi.test;

import com.qtdbp.poi.excel.ExcelReaderUtil;
import com.qtdbp.poi.excel.service.ExcelReaderService;
import com.qtdbp.poi.excel.service.IExcelReader;

/**
 * @author: caidchen
 * @create: 2017-05-23 8:41
 * To change this template use File | Settings | File Templates.
 */
public class Main {

    public static void main(String[] args) throws Exception {

        String file = "D:\\tmp\\物流设施.xlsx";

        IExcelReader reader = new ExcelReaderService();
        //ExcelReaderUtil.readExcel(reader, "F://te03.xls");
        ExcelReaderUtil.readExcel(reader, file);

    }
}
