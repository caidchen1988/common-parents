package com.qtdbp.poi.excel.model;

import java.io.Serializable;
import java.util.List;

/**
 * 解析Excel数据
 *
 * @author: caidchen
 * @create: 2017-05-22 13:40
 * To change this template use File | Settings | File Templates.
 */
public class ExcelSheetModel implements Serializable {

    private String name;    // 名称
    private int totalNum ;      // 总记录数
    private List<ExcelRowModel> rowList ;  // 每行数据

    public ExcelSheetModel(String name, List<ExcelRowModel> rowList) {
        this.name = name;
        this.rowList = rowList;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getTotalNum() {
        return totalNum;
    }

    public void setTotalNum(int totalNum) {
        this.totalNum = totalNum;
    }

    public List<ExcelRowModel> getRowList() {
        return rowList;
    }

    public void setRowList(List<ExcelRowModel> rowList) {
        this.rowList = rowList;
    }

}
