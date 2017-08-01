package com.qtdbp.poi.csv;

import com.csvreader.CsvReader;
import com.qtdbp.poi.excel.model.ExcelCellModel;
import com.qtdbp.poi.excel.model.ExcelRowModel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.List;

/**
 * 利用JavaCSV API来读写csv文件
 *
 * @author: caidchen
 * @create: 2017-07-31 16:37
 * To change this template use File | Settings | File Templates.
 */
public class CSVReader {

    private static final Charset GBK = Charset.forName("gbk") ;
    private static final char DEFAULT_FORMAT = ',' ;

    public static List<ExcelRowModel> csv2Bean(InputStream inputStream) throws IOException {
        return csv2Bean(inputStream, DEFAULT_FORMAT, GBK) ;
    }

    public static List<ExcelRowModel> csv2Bean(InputStream inputStream, char format, Charset charset) throws IOException {

        List<ExcelRowModel> rowList = new ArrayList<>() ;
        // 创建CSV读对象
        CsvReader csvReader = new CsvReader(inputStream, format, charset);

        // 读表头
        csvReader.readHeaders();
        while (csvReader.readRecord()){
            // 读一整行
            ExcelRowModel rowModel = parseRawRecord((int) csvReader.getCurrentRecord(), csvReader.getValues()) ;

            rowList.add(rowModel);
        }

        return rowList ;
    }

    private static ExcelRowModel parseRawRecord(int rowIndex, String[] rawRecords) {

        if(rawRecords == null) return null ;
        List<ExcelCellModel> cellList = new ArrayList<>() ;

        for (int i=0;i<rawRecords.length;i++) {
            cellList.add(new ExcelCellModel(i, rawRecords[i]));

            // 读这行的每一列
//            System.out.print("  ["+i+", "+rawRecords[i]+"]");
        }

        return new ExcelRowModel(rowIndex,cellList) ;
    }

    public static void main(String[] args) throws Exception {

        /*try {
            // 创建CSV读对象
            CsvReader csvReader = new CsvReader("D:\\tmp\\zip\\20882217277885030156_20170725_业务明细.csv", DEFAULT_FORMAT, GBK);

            // 读表头
            csvReader.readHeaders();
            while (csvReader.readRecord()){
                // 读一整行
                System.out.println("["+csvReader.getCurrentRecord()+"]"+csvReader.getValues());

                // 读这行的某一列
//                System.out.println(csvReader.get("Link"));
            }

        } catch (IOException e) {
            e.printStackTrace();
        }*/

        CSVReader.csv2Bean(new FileInputStream(new File("D:\\tmp\\zip\\20882217277885030156_20170725_业务明细.csv"))) ;
    }
}
