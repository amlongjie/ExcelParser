package com.snakotech.excelhelper.handler.impl;

import com.snakotech.excelhelper.IParserParam;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.SAXHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class Excel2007ParseHandler<T> extends BaseExcelParseHandler<T> implements XSSFSheetXMLHandler.SheetContentsHandler {

    private int currentRow = -1;
    private int currentCol = -1;
    private IParserParam parserParam;
    private List<T> result;
    private List<String> rowData;

    public List<T> process(IParserParam parserParam) throws Exception {
        this.parserParam = parserParam;
        result = new ArrayList<>();
        rowData = initRowList(parserParam.getColumnSize());
        OPCPackage xlsxPackage = OPCPackage.open(parserParam.getExcelInputStream());
        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(xlsxPackage, false);
        XSSFReader xssfReader = new XSSFReader(xlsxPackage);
        StylesTable styles = xssfReader.getStylesTable();
        XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        int needParse = 0;
        while (iter.hasNext()) {
            InputStream stream = iter.next();
            if (needParse++ == parserParam.getSheetNum()) {
                processSheet(styles, strings, this, stream);
            }
            stream.close();
        }
        return result;
    }

    private void processSheet(
            StylesTable styles,
            ReadOnlySharedStringsTable strings,
            XSSFSheetXMLHandler.SheetContentsHandler sheetHandler,
            InputStream sheetInputStream) throws IOException, SAXException {
        DataFormatter formatter = new DataFormatter();
        InputSource sheetSource = new InputSource(sheetInputStream);
        try {
            XMLReader sheetParser = SAXHelper.newXMLReader();
            ContentHandler handler = new XSSFSheetXMLHandler(
                    styles, null, strings, sheetHandler, formatter, false);
            sheetParser.setContentHandler(handler);
            sheetParser.parse(sheetSource);
        } catch (ParserConfigurationException e) {
            throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage());
        }
    }

    @Override
    public void startRow(int rowNum) {
        currentRow = rowNum;
        currentCol = -1;
    }

    @Override
    public void endRow(int rowNum) {
        handleEndOfRow(parserParam, rowData, result);
        rowData = initRowList(parserParam.getColumnSize());
        currentCol = -1;
    }




    @Override
    public void cell(String cellReference, String formattedValue,
                     XSSFComment comment) {

        if (cellReference == null) {
            cellReference = new CellAddress(currentRow, currentCol).formatAsString();
        }

        currentCol = (new CellReference(cellReference)).getCol();
        if (currentCol < parserParam.getColumnSize())
            rowData.set(currentCol, formattedValue.trim());
    }

    @Override
    public void headerFooter(String text, boolean isHeader, String tagName) {

    }
}
