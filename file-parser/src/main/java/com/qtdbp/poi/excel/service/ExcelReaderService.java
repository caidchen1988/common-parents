package com.qtdbp.poi.excel.service;

import com.qtdbp.poi.excel.model.ExcelCellModel;
import com.qtdbp.poi.excel.model.ExcelRowModel;
import com.qtdbp.poi.excel.model.ExcelSheetModel;

/**
 * @author: caidchen
 * @create: 2017-05-23 8:41
 * To change this template use File | Settings | File Templates.
 */
public class ExcelReaderService implements IExcelReader {

    public void getSheet(ExcelSheetModel sheet) {

        System.out.println("............................");
        long start = System.currentTimeMillis();

        /*// 方式一：构建excel2007写入器,通过文件流的方式处理
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        try {
            AbstractExcel2007Writer excel07Writer = new Excel2007Writer();
            excel07Writer.process(outStream, sheet);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if(outStream == null) outStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }*/

        /*// 方式二：构建excel2007写入器,通过本地文件方式
        try {
            AbstractExcel2007Writer excel07Writer = new Excel2007Writer();
            excel07Writer.process("D:\\tmp\\"+sheet.getName()+".xlsx", sheet);
        } catch (Exception e) {
            e.printStackTrace();
        }*/

        System.out.println("sheet name: " + sheet.getName()+"; ");
        for (ExcelRowModel row : sheet.getRowList()) {

            System.out.print("##row: " + row.getRowNum() + " , data: ");
            for (ExcelCellModel cell : row.getCellList()) {

                System.out.print("num: " + cell.getColNum() + ", val: " + cell.getColVal()+"; ");
            }
            System.out.println();
        }
        System.out.println();


        long end = System.currentTimeMillis();
        System.out.println("....................."+(end-start));

    }
}
