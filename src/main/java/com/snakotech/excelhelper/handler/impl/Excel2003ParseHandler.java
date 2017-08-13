package com.snakotech.excelhelper.handler.impl;

import com.snakotech.excelhelper.IParserParam;
import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
import org.apache.poi.hssf.model.HSSFFormulaParser;
import org.apache.poi.hssf.record.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

public class Excel2003ParseHandler<T> extends BaseExcelParseHandler<T> implements HSSFListener {

    /**
     * Should we output the formula, or the value it has?
     */
    private boolean outputFormulaValues = true;

    /**
     * For parsing Formulas
     */
    private EventWorkbookBuilder.SheetRecordCollectingListener workbookBuildingListener;
    private HSSFWorkbook stubWorkbook;

    // Records we pick up as we process
    private SSTRecord sstRecord;
    private FormatTrackingHSSFListener formatListener;

    /**
     * So we known which sheet we're on
     */
    private BoundSheetRecord[] orderedBSRs;
    private List<BoundSheetRecord> boundSheetRecords = new ArrayList<BoundSheetRecord>();

    private int nextColumn;
    private boolean outputNextStringRecord;
    private Integer curSheet = -1;


    private IParserParam parserParam;
    private List<String> rowData;
    private List<T> result;


    public List<T> process(IParserParam parserParam) throws Exception {
        this.parserParam = parserParam;
        result = new ArrayList<>();
        rowData = initRowList(parserParam.getColumnSize());
        POIFSFileSystem fs = new POIFSFileSystem(parserParam.getExcelInputStream());
        MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(this);
        formatListener = new FormatTrackingHSSFListener(listener);

        HSSFEventFactory factory = new HSSFEventFactory();
        HSSFRequest request = new HSSFRequest();

        if (outputFormulaValues) {
            request.addListenerForAllRecords(formatListener);
        } else {
            workbookBuildingListener = new EventWorkbookBuilder.SheetRecordCollectingListener(formatListener);
            request.addListenerForAllRecords(workbookBuildingListener);
        }

        factory.processWorkbookEvents(request, fs);
        return result;
    }


    @Override
    public void processRecord(Record record) {

        int thisColumn = -1;
        String thisStr = null;

        switch (record.getSid()) {
            case BoundSheetRecord.sid:
                boundSheetRecords.add((BoundSheetRecord) record);
                break;
            case BOFRecord.sid:
                BOFRecord br = (BOFRecord) record;
                if (br.getType() == BOFRecord.TYPE_WORKSHEET) {
                    // Create sub workbook if required
                    if (workbookBuildingListener != null && stubWorkbook == null) {
                        stubWorkbook = workbookBuildingListener.getStubHSSFWorkbook();
                    }

                    // Output the worksheet name
                    // Works by ordering the BSRs by the location of
                    //  their BOFRecords, and then knowing that we
                    //  process BOFRecords in byte offset order
                    curSheet++;
                    if (orderedBSRs == null) {
                        orderedBSRs = BoundSheetRecord.orderByBofPosition(boundSheetRecords);
                    }
                }
                break;

            case SSTRecord.sid:
                sstRecord = (SSTRecord) record;
                break;

            case BlankRecord.sid:
                BlankRecord brec = (BlankRecord) record;

                thisColumn = brec.getColumn();
                thisStr = "";
                break;
            case BoolErrRecord.sid:
                BoolErrRecord berec = (BoolErrRecord) record;
                thisColumn = berec.getColumn();
                thisStr = "";
                break;

            case FormulaRecord.sid:
                FormulaRecord frec = (FormulaRecord) record;
                thisColumn = frec.getColumn();

                if (outputFormulaValues) {
                    if (Double.isNaN(frec.getValue())) {
                        // Formula result is a string
                        // This is stored in the next record
                        outputNextStringRecord = true;
                        nextColumn = frec.getColumn();
                    } else {
                        thisStr = formatListener.formatNumberDateCell(frec);
                    }
                } else {
                    thisStr =
                            HSSFFormulaParser.toFormulaString(stubWorkbook, frec.getParsedExpression());
                }
                break;
            case StringRecord.sid:
                if (outputNextStringRecord) {
                    // String for formula
                    StringRecord srec = (StringRecord) record;
                    thisStr = srec.getString();
                    thisColumn = nextColumn;
                    outputNextStringRecord = false;
                }
                break;

            case LabelRecord.sid:
                LabelRecord lrec = (LabelRecord) record;

                thisColumn = lrec.getColumn();
                thisStr = lrec.getValue();
                break;
            case LabelSSTRecord.sid:
                LabelSSTRecord lsrec = (LabelSSTRecord) record;

                thisColumn = lsrec.getColumn();
                if (sstRecord == null) {
//                    thisStr = "(No SST Record, can't identify string)";
                    thisStr = "";
                } else {
                    thisStr = sstRecord.getString(lsrec.getSSTIndex()).toString();
                }
                break;
            case NoteRecord.sid:
                NoteRecord nrec = (NoteRecord) record;
                thisColumn = nrec.getColumn();
                thisStr = "";
                break;
            case NumberRecord.sid:
                NumberRecord numrec = (NumberRecord) record;

                thisColumn = numrec.getColumn();

                // Format
                thisStr = formatListener.formatNumberDateCell(numrec);
                break;
            case RKRecord.sid:
                RKRecord rkrec = (RKRecord) record;

                thisColumn = rkrec.getColumn();
                thisStr = "";
                break;
            default:
                break;
        }
        if (!Objects.equals(curSheet, parserParam.getSheetNum())) {
            return;
        }

        // Handle missing column
        if (record instanceof MissingCellDummyRecord) {
            MissingCellDummyRecord mc = (MissingCellDummyRecord) record;
            thisColumn = mc.getColumn();
            thisStr = "";
        }

        // If we got something to print out, do so
        if (thisStr != null && thisColumn < parserParam.getColumnSize()) {
            rowData.set(thisColumn, thisStr.trim());
        }

        // Handle end of row
        if (record instanceof LastCellOfRowDummyRecord) {
            // End the row
            handleEndOfRow(parserParam, rowData, result);
            rowData = initRowList(parserParam.getColumnSize());
        }
    }
}
