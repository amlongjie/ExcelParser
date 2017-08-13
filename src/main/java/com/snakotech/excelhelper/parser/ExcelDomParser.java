package com.snakotech.excelhelper.parser;

import com.snakotech.excelhelper.handler.impl.ExcelDomParseHandler;
import com.snakotech.excelhelper.handler.IExcelParseHandler;

import java.io.InputStream;

public class ExcelDomParser<T> extends AbstractExcelParser<T> {

    private IExcelParseHandler<T> excelParseHandler;

    public ExcelDomParser() {
        this.excelParseHandler = new ExcelDomParseHandler<>();
    }

    @Override
    protected IExcelParseHandler<T> createHandler(InputStream excelInputStream) {
        return this.excelParseHandler;
    }
}
