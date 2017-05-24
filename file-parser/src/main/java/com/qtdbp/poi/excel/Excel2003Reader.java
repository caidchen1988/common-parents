package com.qtdbp.poi.excel;

import com.qtdbp.poi.excel.model.ExcelCellModel;
import com.qtdbp.poi.excel.service.IExcelReader;
import org.apache.poi.hssf.eventusermodel.EventWorkbookBuilder.SheetRecordCollectingListener;
import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
import org.apache.poi.hssf.model.HSSFFormulaParser;
import org.apache.poi.hssf.record.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * 抽象Excel2003读取器，通过实现HSSFListener监听器，采用事件驱动模式解析excel2003 中的内容，遇到特定事件才会触发，大大减少了内存的使用。
 *
 * @author: caidchen
 * @create: 2017-05-23 8:36
 * To change this template use File | Settings | File Templates.
 */
public class Excel2003Reader  implements HSSFListener {
    private int minColumns = -1;
    private POIFSFileSystem fs;
    private int lastRowNumber;
    private int lastColumnNumber;

    /**
     * Should we output the formula, or the value it has?
     */
    private boolean outputFormulaValues = true;

    /**
     * For parsing Formulas
     */
    private SheetRecordCollectingListener workbookBuildingListener;
    //excel2003工作薄
    private HSSFWorkbook stubWorkbook;

    // Records we pick up as we process
    private SSTRecord sstRecord;
    private FormatTrackingHSSFListener formatListener;

    //表索引
    private int sheetIndex = -1;
    private BoundSheetRecord[] orderedBSRs;
    @SuppressWarnings("unchecked")
    private ArrayList boundSheetRecords = new ArrayList();

    // For handling formulas with string results
    private int nextRow;
    private int nextColumn;
    private boolean outputNextStringRecord;
    //当前行
    private int curRow = 0;
    //存储行记录的容器
    private List<ExcelCellModel> cellList = new ArrayList<ExcelCellModel>();
    ;
    @SuppressWarnings("unused")
    private String sheetName;

    private IExcelReader rowReader;


    public void setRowReader(IExcelReader rowReader) {
        this.rowReader = rowReader;
    }

    /**
     * 遍历excel下所有的sheet
     *
     * @param localFileName 本地文件
     * @throws IOException
     */
    public void process(String localFileName) throws IOException {
        process(new FileInputStream(localFileName)) ;
    }

    /**
     * 遍历excel下所有的sheet
     *
     * @param stream 文件流
     * @throws IOException
     */
    public void process(InputStream stream) throws IOException {
        this.fs = new POIFSFileSystem(stream);
        MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(
                this);
        formatListener = new FormatTrackingHSSFListener(listener);
        HSSFEventFactory factory = new HSSFEventFactory();
        HSSFRequest request = new HSSFRequest();
        if (outputFormulaValues) {
            request.addListenerForAllRecords(formatListener);
        } else {
            workbookBuildingListener = new SheetRecordCollectingListener(
                    formatListener);
            request.addListenerForAllRecords(workbookBuildingListener);
        }
        factory.processWorkbookEvents(request, fs);
    }

    /**
     * HSSFListener 监听方法，处理 Record
     */
    @SuppressWarnings("unchecked")
    public void processRecord(Record record) {
        int thisRow = -1;
        int thisColumn = -1;
        String thisStr = null;
        String value = null;
        switch (record.getSid()) {
            case BoundSheetRecord.sid:
                boundSheetRecords.add(record);
                break;
            case BOFRecord.sid:
                BOFRecord br = (BOFRecord) record;
                if (br.getType() == BOFRecord.TYPE_WORKSHEET) {
                    // 如果有需要，则建立子工作薄
                    if (workbookBuildingListener != null && stubWorkbook == null) {
                        stubWorkbook = workbookBuildingListener
                                .getStubHSSFWorkbook();
                    }

                    sheetIndex++;
                    if (orderedBSRs == null) {
                        orderedBSRs = BoundSheetRecord
                                .orderByBofPosition(boundSheetRecords);
                    }
                    sheetName = orderedBSRs[sheetIndex].getSheetname();
                }
                break;

            case SSTRecord.sid:
                sstRecord = (SSTRecord) record;
                break;

            case BlankRecord.sid:
                BlankRecord brec = (BlankRecord) record;
                thisRow = brec.getRow();
                thisColumn = brec.getColumn();
                thisStr = "";
                cellList.add(new ExcelCellModel(thisColumn, thisStr));
                break;
            case BoolErrRecord.sid: //单元格为布尔类型
                BoolErrRecord berec = (BoolErrRecord) record;
                thisRow = berec.getRow();
                thisColumn = berec.getColumn();
                thisStr = berec.getBooleanValue() + "";
                cellList.add(new ExcelCellModel(thisColumn, thisStr));
                break;

            case FormulaRecord.sid: //单元格为公式类型
                FormulaRecord frec = (FormulaRecord) record;
                thisRow = frec.getRow();
                thisColumn = frec.getColumn();
                if (outputFormulaValues) {
                    if (Double.isNaN(frec.getValue())) {
                        // Formula result is a string
                        // This is stored in the next record
                        outputNextStringRecord = true;
                        nextRow = frec.getRow();
                        nextColumn = frec.getColumn();
                    } else {
                        thisStr = formatListener.formatNumberDateCell(frec);
                    }
                } else {
                    thisStr = '"' + HSSFFormulaParser.toFormulaString(stubWorkbook,
                            frec.getParsedExpression()) + '"';
                }
                cellList.add(new ExcelCellModel(thisColumn, thisStr));
                break;
            case StringRecord.sid://单元格中公式的字符串
                if (outputNextStringRecord) {
                    // String for formula
                    StringRecord srec = (StringRecord) record;
                    thisStr = srec.getString();
                    thisRow = nextRow;
                    thisColumn = nextColumn;
                    outputNextStringRecord = false;
                }
                break;
            case LabelRecord.sid:
                LabelRecord lrec = (LabelRecord) record;
                curRow = thisRow = lrec.getRow();
                thisColumn = lrec.getColumn();
                value = lrec.getValue().trim();
                value = value.equals("") ? " " : value;
                cellList.add(new ExcelCellModel(thisColumn, value));
                break;
            case LabelSSTRecord.sid:  //单元格为字符串类型
                LabelSSTRecord lsrec = (LabelSSTRecord) record;
                curRow = thisRow = lsrec.getRow();
                thisColumn = lsrec.getColumn();
                if (sstRecord == null) {
                    cellList.add(new ExcelCellModel(thisColumn, " "));
                } else {
                    value = sstRecord
                            .getString(lsrec.getSSTIndex()).toString().trim();
                    value = value.equals("") ? " " : value;
                    cellList.add(new ExcelCellModel(thisColumn, value));
                }
                break;
            case NumberRecord.sid:  //单元格为数字类型
                NumberRecord numrec = (NumberRecord) record;
                curRow = thisRow = numrec.getRow();
                thisColumn = numrec.getColumn();
                value = formatListener.formatNumberDateCell(numrec).trim();
                value = value.equals("") ? " " : value;
                // 向容器加入列值
                cellList.add(new ExcelCellModel(thisColumn, value));
                break;
            default:
                break;
        }

        // 遇到新行的操作
        if (thisRow != -1 && thisRow != lastRowNumber) {
            lastColumnNumber = -1;
        }

        // 空值的操作
        if (record instanceof MissingCellDummyRecord) {
            MissingCellDummyRecord mc = (MissingCellDummyRecord) record;
            curRow = thisRow = mc.getRow();
            thisColumn = mc.getColumn();
            cellList.add(new ExcelCellModel(thisColumn, " "));
        }

        // 更新行和列的值
        if (thisRow > -1)
            lastRowNumber = thisRow;
        if (thisColumn > -1)
            lastColumnNumber = thisColumn;

        // 行结束时的操作
        if (record instanceof LastCellOfRowDummyRecord) {
            if (minColumns > 0) {
                // 列值重新置空
                if (lastColumnNumber == -1) {
                    lastColumnNumber = 0;
                }
            }
            lastColumnNumber = -1;
            // 每行结束时， 调用getRows() 方法
//            rowReader.getRows(sheetIndex, curRow, cellList);

            // 清空容器
            cellList.clear();
        }
    }
}