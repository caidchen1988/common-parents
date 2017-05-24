package com.qtdbp.poi.excel;

import com.qtdbp.poi.excel.model.ExcelCellModel;
import com.qtdbp.poi.excel.model.ExcelRowModel;
import com.qtdbp.poi.excel.model.ExcelSheetModel;

import java.util.List;

/**
 * @author: caidchen
 * @create: 2017-05-23 10:36
 * To change this template use File | Settings | File Templates.
 */
public class Excel2007Writer extends AbstractExcel2007Writer {

    /*
     * 可根据需求重写此方法，对于单元格的小数或者日期格式，会出现精度问题或者日期格式转化问题，建议使用字符串插入方法
     * @see com.excel.ver2.AbstractExcel2007Writer#generate()
     */
    @Override
    public void generate() throws Exception {
        //电子表格开始
        beginSheet();

        for (int rownum = 0; rownum < 1000000; rownum++) {
            //插入新行
            insertRow(rownum);
            //建立新单元格,索引值从0开始,表示第一列
            createCell(0, "中国<" + rownum + "!");
            createCell(1, 34343.123456789);
            createCell(2, "23.67%");
            createCell(3, "12:12:23");
            createCell(4, "2010-10-11 12:12:23");
            createCell(5, "true");
            createCell(6, "false");

            //结束行
            endRow();
        }
        //电子表格结束
        endSheet();
    }

    @Override
    public void generate(ExcelSheetModel sheetModel) throws Exception {
        //电子表格开始
        beginSheet();
        List<ExcelRowModel> rowList = sheetModel.getRowList() ;
        if(rowList != null && !rowList.isEmpty()) {
            for (ExcelRowModel row : rowList) {
                //插入新行
                insertRow(row.getRowNum());
                //建立新单元格,索引值从0开始,表示第一列
                List<ExcelCellModel> cellList = row.getCellList() ;
                if(cellList != null) {
                    for (ExcelCellModel cell : cellList)
                    createCell(cell.getColNum(), cell.getColVal());
                }
                //结束行
                endRow();
            }
        }
        //电子表格结束
        endSheet();

    }


    /**
     * @param args
     * @throws Exception
     */
    public static void main(String[] args) throws Exception {
        // TODO Auto-generated method stub
        System.out.println("............................");
        long start = System.currentTimeMillis();
        //构建excel2007写入器
        AbstractExcel2007Writer excel07Writer = new Excel2007Writer();
        //调用处理方法
        excel07Writer.process("D:\\tmp\\test07.xlsx");
        long end = System.currentTimeMillis();
        System.out.println("....................."+(end-start)/1000);
    }
}
