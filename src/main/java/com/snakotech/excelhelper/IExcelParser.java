package com.snakotech.excelhelper;

import java.util.List;

public interface IExcelParser<T> {
    List<T> parse(IParserParam parserParam);
}
