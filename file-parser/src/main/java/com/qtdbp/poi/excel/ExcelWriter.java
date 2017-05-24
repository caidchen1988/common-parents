package com.qtdbp.poi.excel;

import com.qtdbp.poi.excel.model.ExcelCellModel;
import com.qtdbp.poi.excel.model.ExcelRowModel;
import com.qtdbp.poi.excel.model.ExcelSheetModel;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

/**
 * 基于POI生成Excel
 *
 * @author: caidchen
 * @create: 2017-05-22 15:13
 * To change this template use File | Settings | File Templates.
 */
@Deprecated
public class ExcelWriter<T> {

    /**
     * 导出excel
     * @param title
     * @param rowList
     * @return
     */
    public static Workbook exportExcel(String title, List<ExcelRowModel> rowList) {

        if(rowList == null) return null ;

        Workbook workbook = new SXSSFWorkbook(500);

        CellStyle style = setCellStyle(workbook);
        // 生成一个表格
        Sheet sheet = workbook.createSheet(title);
        // 设置表格默认列宽度为15个字节
        sheet.setDefaultColumnWidth((short) 15);

        // 产生表格标题行
        Row row = sheet.createRow(0);

        ExcelRowModel headerModel = rowList.get(0) ;
        List<ExcelCellModel> headers = headerModel.getCellList() ;

        for (ExcelCellModel header : headers) {
            Cell cell = row.createCell(header.getColNum());
            cell.setCellStyle(style);
            HSSFRichTextString text = new HSSFRichTextString(header.getColVal());
            cell.setCellValue(text);
        }

        if(rowList.size() > 1) {
            rowList.remove(0);// 移除第一个元素

            int index = 0;
            for (ExcelRowModel rowModel : rowList) {
                index++;
                try {
                    row = sheet.createRow(index);

                    List<ExcelCellModel> cellList = rowModel.getCellList();
                    if (cellList != null && !cellList.isEmpty()) {

                        for (ExcelCellModel cellModel : cellList) {
                            Cell cell = row.createCell(cellModel.getColNum());

                            if (cellModel.getColVal() == null) continue;
                            HSSFRichTextString richString = new HSSFRichTextString(cellModel.getColVal());
                            /*HSSFFont font3 = workbook.createFont();
                            font3.setColor(HSSFColor.BLUE.index);
                            richString.applyFont(font3);*/
                            cell.setCellValue(richString);
                        }
                    }
                } catch (SecurityException ex) {
                    ex.getMessage();
                } catch (IllegalArgumentException ex) {
                    ex.getMessage();
                }

            }
        }

        return workbook ;
    }

    /**
     * 设置工作表样式
     * @param workbook
     * @return
     */
    protected static CellStyle setCellStyle(Workbook workbook) {
        // 生成一个样式
        CellStyle style = workbook.createCellStyle();
        // 设置这些样式
        style.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);

        // 生成一个字体
        Font font = workbook.createFont();
        font.setColor(HSSFColor.VIOLET.index);
        font.setFontHeightInPoints((short) 12);
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);

        // 把字体应用到当前的样式
        style.setFont(font);
        // 生成并设置另一个样式
        CellStyle style2 = workbook.createCellStyle();
        style2.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
        style2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style2.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style2.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style2.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style2.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        // 生成另一个字体
        Font font2 = workbook.createFont();
        font2.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
        // 把字体应用到当前的样式
        style2.setFont(font2);

        return style ;
    }

    @Test
    public void write() {

        // 解析excel, 拆分sheet
        List<ExcelSheetModel> list = null;
        try {
            list = ExcelReader.readExcel("C:\\Users\\dell\\Desktop\\物流(2)\\全国物流企业.xlsx");
        } catch (Exception e) {
            e.printStackTrace();
        }

        // 基于每个sheet,生成excel
        if(list != null) {
            for (ExcelSheetModel sheetModel : list) {
                FileOutputStream os = null ;
                try {
                    Workbook workbook = exportExcel(sheetModel.getName(), sheetModel.getRowList());

//                    os = new FileOutputStream(new File("D:\\tmp\\" + sheetModel.getName() + "."+sheetModel.getExtension()));
                    workbook.write(os);

                } catch (Exception e) {
                    e.printStackTrace();
                } finally {
                    try {
                       if(os != null) os.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }
        }
    }

}
