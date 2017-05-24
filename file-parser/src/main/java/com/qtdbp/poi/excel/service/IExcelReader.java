package com.qtdbp.poi.excel.service;

import com.qtdbp.poi.excel.model.ExcelCellModel;
import com.qtdbp.poi.excel.model.ExcelRowModel;
import com.qtdbp.poi.excel.model.ExcelSheetModel;

import java.util.List;

/**
 * @author: caidchen
 * @create: 2017-05-23 8:40
 * To change this template use File | Settings | File Templates.
 */
public interface IExcelReader {

    /**业务逻辑实现方法
     * @param sheetIndex
     * @param curRow
     * @param rowlist
     */
//   ExcelRowModel getRows(int sheetIndex, int curRow, List<ExcelCellModel> rowlist);

    /**
     * 解析每个sheet表格数据，处理业务逻辑
     * @param sheet
     */
    void getSheet(ExcelSheetModel sheet);
}
