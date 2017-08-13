package com.snakotech.excelhelper.parser;


import com.snakotech.excelhelper.handler.IExcelParseHandler;
import com.snakotech.excelhelper.handler.impl.Excel2003ParseHandler;
import com.snakotech.excelhelper.handler.impl.Excel2007ParseHandler;
import org.apache.poi.poifs.filesystem.DocumentFactoryHelper;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.util.IOUtils;

import java.io.InputStream;

public class ExcelSaxParser<T> extends AbstractExcelParser<T> {

    public IExcelParseHandler<T> createHandler(InputStream excelInputStream) {
        try {
            byte[] header8 = IOUtils.peekFirst8Bytes(excelInputStream);
            if (NPOIFSFileSystem.hasPOIFSHeader(header8)) {
                return new Excel2003ParseHandler<>();
            } else if (DocumentFactoryHelper.hasOOXMLHeader(excelInputStream)) {
                return new Excel2007ParseHandler<>();
            } else {
                throw new IllegalArgumentException("Your InputStream was neither an OLE2 stream, nor an OOXML stream");
            }
        } catch (Exception e) {
            logger.error("getParserInstance Error!", e);
            throw new RuntimeException(e);
        }
    }

}
