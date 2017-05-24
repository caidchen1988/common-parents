package com.qtdbp.poi.excel.model;

/**
 * Excel每列数据
 *
 * @author: caidchen
 * @create: 2017-05-22 14:49
 * To change this template use File | Settings | File Templates.
 */
public class ExcelCellModel {

    private int colNum ; // 列数
    private String colVal ; // 数据

    public ExcelCellModel(int colNum, String colVal) {
        this.colNum = colNum;
        this.colVal = colVal;
    }

    public int getColNum() {
        return colNum;
    }

    public void setColNum(int colNum) {
        this.colNum = colNum;
    }

    public String getColVal() {
        return colVal;
    }

    public void setColVal(String colVal) {
        this.colVal = colVal;
    }
}
