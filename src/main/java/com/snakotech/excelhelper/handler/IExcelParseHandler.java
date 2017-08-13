package com.snakotech.excelhelper.handler;

import com.snakotech.excelhelper.IParserParam;

import java.util.List;

public interface IExcelParseHandler<T> {

    List<T> process(IParserParam parserParam) throws Exception;

}
