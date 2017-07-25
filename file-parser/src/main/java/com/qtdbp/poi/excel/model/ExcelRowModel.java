package com.qtdbp.poi.excel.model;

import java.util.List;

/**
 * Excel每行数据
 *
 * @author: caidchen
 * @create: 2017-05-22 14:49
 * To change this template use File | Settings | File Templates.
 */
public class ExcelRowModel {

    private int rowNum ; // 行数
    private int totalCellNum ; // 一行总列数
    private List<ExcelCellModel> cellList ; // 当前所有列数据

    public ExcelRowModel(int rowNum) {
        this.rowNum = rowNum;
    }

    public ExcelRowModel(int rowNum, List<ExcelCellModel> cellList) {
        this.rowNum = rowNum;
        this.cellList = cellList;
    }

    public ExcelRowModel(int rowNum, List<ExcelCellModel> cellList, int totalCellNum) {
        this.rowNum = rowNum;
        this.cellList = cellList;
        this.totalCellNum = totalCellNum;
    }

    public int getRowNum() {
        return rowNum;
    }

    public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
    }

    public List<ExcelCellModel> getCellList() {
        return cellList;
    }

    public void setCellList(List<ExcelCellModel> cellList) {
        this.cellList = cellList;
    }

    public int getTotalCellNum() {
        return totalCellNum;
    }

    public void setTotalCellNum(int totalCellNum) {
        this.totalCellNum = totalCellNum;
    }
}
