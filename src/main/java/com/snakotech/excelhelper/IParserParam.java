package com.snakotech.excelhelper;

import java.io.InputStream;
import java.util.List;

public interface IParserParam {

    Integer FIRST_SHEET = 0;

    InputStream getExcelInputStream();

    Class getTargetClass();

    Integer getColumnSize();

    Integer getSheetNum();

    List<String> getHeader();
}
