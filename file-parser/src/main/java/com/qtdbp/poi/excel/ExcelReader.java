package com.qtdbp.poi.excel;

import com.qtdbp.poi.excel.model.ExcelCellModel;
import com.qtdbp.poi.excel.model.ExcelRowModel;
import com.qtdbp.poi.excel.model.ExcelSheetModel;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

/**
 * 基于POI操作Excel表格的功能类
 *
 * @author: caidchen
 * @create: 2017-05-22 11:29
 * To change this template use File | Settings | File Templates.
 */
@Deprecated
public class ExcelReader {


    protected static final String dateTimeFmtPattern = "yyyy-MM-dd HH:mm:ss";

    protected static final String dateFmtPattern = "yyyy-MM-dd";

    /**
     * 读取excel文件（同时支持2003和2007格式）
     *
     * @param fileName
     *            文件名，绝对路径
     * @return list中的map的key是列的序号
     * @throws Exception
     *             io异常等
     */
    public static List<ExcelSheetModel> readExcel(String fileName) throws Exception {
        FileInputStream fis = null;
        Workbook wb = null;
        List<ExcelSheetModel> list = null;
        try {
            String extension = FilenameUtils.getExtension(fileName);

            fis = new FileInputStream(fileName);
            list = read(fis, extension);

            return list;

        } finally {
            if (null != wb) {
                wb.close();
            }

            if (null != fis) {
                fis.close();
            }
        }

    }

    /**
     * 读取excel文件（同时支持2003和2007格式）
     *
     * @param fis
     *            文件输入流
     * @param extension
     *            文件名扩展名: xls 或 xlsx 不区分大小写
     * @return list中的map的key是列的序号
     * @throws Exception
     *             io异常等
     */
    public static List<ExcelSheetModel> read(FileInputStream fis, String extension) throws Exception {

        Workbook wb = null;
        List<ExcelSheetModel> list = null;
        try {

            if ("xls".equalsIgnoreCase(extension)) {
                wb = new HSSFWorkbook(fis);
            } else if ("xlsx".equalsIgnoreCase(extension)) {
                wb = new XSSFWorkbook(fis);
            } else {
                throw new Exception("file is not office excel");
            }

            list = readWorkbook(wb,extension);

            return list;

        } finally {
            if (null != wb) {
                wb.close();
            }
        }

    }

    /**
     * 读取每个工作簿
     * @param wb
     * @return
     * @throws Exception
     */
    protected static List<ExcelSheetModel> readWorkbook(Workbook wb, String extension) throws Exception {

        List<ExcelSheetModel> list = new LinkedList<ExcelSheetModel>();

        // 解析每个sheet
        for (int k = 0; k < wb.getNumberOfSheets(); k++) {
            Sheet sheet = wb.getSheetAt(k);
            int rows = sheet.getPhysicalNumberOfRows();

            ExcelSheetModel sheetModel = null ;//new ExcelSheetModel(sheet.getSheetName(), rows) ;
//            sheetModel.setExtension(extension);

            List<ExcelRowModel> rowList = null ;
            if(rows > 0) rowList = new LinkedList<ExcelRowModel>();
            // 解析每行
            int headerCells = 0 ; // 记录第一行列数
            for (int r = 0; r < rows; r++) {
                ExcelRowModel rowModel = new ExcelRowModel(r) ;

                Row row = sheet.getRow(r);
                if (row == null) continue;
                int cells = row.getPhysicalNumberOfCells();
                if(r == 0) headerCells = cells ;

                List<ExcelCellModel> cellList = new LinkedList<ExcelCellModel>();
                // 解析每列
                for (int c = 0; c < headerCells; c++) {
                    Cell cell = row.getCell(c);

                    if (cell == null) continue;

                    cellList.add(new ExcelCellModel(cell.getColumnIndex(), getCellValue(cell)));
                }
                rowModel.setCellList(cellList);

                rowList.add(rowModel);
            }
            sheetModel.setRowList(rowList);

            list.add(sheetModel);
        }

        return list;
    }

    /**
     * 解析每个单元格数据
     * @param cell
     * @return
     */
    protected static String getCellValue(Cell cell) {
        String value = null;

        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_FORMULA: // 公式
            case Cell.CELL_TYPE_NUMERIC: // 数字

                double doubleVal = cell.getNumericCellValue();
                short format = cell.getCellStyle().getDataFormat();
                String formatString = cell.getCellStyle().getDataFormatString();

                if (format == 14 || format == 31 || format == 57 || format == 58 || (format >= 176 && format <= 183)) {
                    // 日期
                    Date date = DateUtil.getJavaDate(doubleVal);
                    value = formatDate(date, dateFmtPattern);
                } else if (format == 20 || format == 32 || (format >= 184 && format <= 187)) {
                    // 时间
                    Date date = DateUtil.getJavaDate(doubleVal);
                    value = formatDate(date, "HH:mm");
                } else {
                    value = String.valueOf(doubleVal);
                }

                break;
            case Cell.CELL_TYPE_STRING: // 字符串
                value = cell.getStringCellValue();

                break;
            case Cell.CELL_TYPE_BLANK: // 空白
                value = "";
                break;
            case Cell.CELL_TYPE_BOOLEAN: // Boolean
                value = String.valueOf(cell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_ERROR: // Error，返回错误码
                value = String.valueOf(cell.getErrorCellValue());
                break;
            default:
                value = "";
                break;
        }
        return value;
    }

    /**
     * 时间格式转换
     * @param d
     * @param sdf
     * @return
     */
    @SuppressWarnings("deprecation")
    private static String formatDate(Date d, String sdf) {
        String value = null;

        SimpleDateFormat dateFormat = null ;
        if (d.getSeconds() == 0 && d.getMinutes() == 0 && d.getHours() == 0) {
            dateFormat = new SimpleDateFormat(dateTimeFmtPattern);
        } else {
            dateFormat = new SimpleDateFormat(sdf);
        }
        value = dateFormat.format(d);

        return value;
    }


    @Test
    public void testReader() throws Exception {
        List<ExcelSheetModel> list = readExcel("C:\\Users\\dell\\Desktop\\物流(2)\\EDI指数.xlsx");

        if(list != null) {
            for (ExcelSheetModel sheetModel : list) {
                System.out.println("########################sheet: " + sheetModel.getName());
                for (ExcelRowModel row : sheetModel.getRowList()) {

                    System.out.print("##row: " + row.getRowNum() + " , data: ");
                    for (ExcelCellModel cell : row.getCellList()) {

                        System.out.print("num: " + cell.getColNum() + ", val: " + cell.getColVal()+"; ");
                    }
                    System.out.println();
                }
                System.out.println();
            }
        }
    }

}
