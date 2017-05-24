package com.qtdbp.poi.excel.service;

import com.qtdbp.poi.excel.model.ExcelSheetModel;

/**
 * @author: caidchen
 * @create: 2017-05-23 8:40
 * To change this template use File | Settings | File Templates.
 */
public interface IExcelReader {
    /**
     * 解析每个sheet表格数据，处理业务逻辑
     * @param sheet
     */
    void getSheet(ExcelSheetModel sheet);
}
