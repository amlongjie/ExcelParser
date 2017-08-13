package com.snakotech;


import com.snakotech.excelhelper.IExcelParser;
import com.snakotech.excelhelper.IParserParam;
import com.snakotech.excelhelper.param.DefaultParserParam;
import com.snakotech.excelhelper.parser.ExcelDomParser;
import com.snakotech.excelhelper.parser.ExcelSaxParser;
import org.junit.Test;

import java.util.List;

import static org.junit.Assert.fail;

public class ParserTest {

    private IExcelParser<User> parser;

    @Test
    public void testDomXlsx() {

        parser = new ExcelDomParser<>();

        IParserParam parserParam = DefaultParserParam.builder()
                .excelInputStream(Thread.currentThread().getContextClassLoader()
                        .getResourceAsStream("test01.xlsx"))
                .columnSize(4)
                .sheetNum(IParserParam.FIRST_SHEET)
                .targetClass(User.class)
                .header(User.getHeader())
                .build();

        List<User> user = parser.parse(parserParam);
        System.out.println(user);
    }

    @Test
    public void testDomXls() {

        parser = new ExcelDomParser<>();

        IParserParam parserParam = DefaultParserParam.builder()
                .excelInputStream(Thread.currentThread().getContextClassLoader().getResourceAsStream("test01.xls"))
                .columnSize(4)
                .sheetNum(IParserParam.FIRST_SHEET)
                .targetClass(User.class)
                .header(User.getHeader())
                .build();

        List<User> user = parser.parse(parserParam);
        System.out.println(user);
    }

    @Test
    public void testSaxXlsx() {
        parser = new ExcelSaxParser<>();

        IParserParam parserParam = DefaultParserParam.builder()
                .excelInputStream(Thread.currentThread().getContextClassLoader().getResourceAsStream("test01.xlsx"))
                .columnSize(4)
                .sheetNum(IParserParam.FIRST_SHEET)
                .targetClass(User.class)
                .header(User.getHeader())
                .build();

        List<User> user = parser.parse(parserParam);
        System.out.println(user);
    }


    @Test
    public void testSaxXls() {
        parser = new ExcelSaxParser<>();

        IParserParam parserParam = DefaultParserParam.builder()
                .excelInputStream(Thread.currentThread().getContextClassLoader().getResourceAsStream("test01.xls"))
                .columnSize(4)
                .sheetNum(IParserParam.FIRST_SHEET)
                .targetClass(User.class)
                .header(User.getHeader())
                .build();

        List<User> user = parser.parse(parserParam);
        System.out.println(user);
    }

    @Test
    public void testSheet02Xls() {
        parser = new ExcelSaxParser<>();

        IParserParam parserParam = DefaultParserParam.builder()
                .excelInputStream(Thread.currentThread().getContextClassLoader().getResourceAsStream("test02.xls"))
                .columnSize(4)
                .sheetNum(IParserParam.FIRST_SHEET + 1)
                .targetClass(User.class)
                .header(User.getHeader())
                .build();

        List<User> user = parser.parse(parserParam);
        System.out.println(user);
    }

    @Test
    public void testSheet02Xlsx() {
        parser = new ExcelSaxParser<>();

        IParserParam parserParam = DefaultParserParam.builder()
                .excelInputStream(Thread.currentThread().getContextClassLoader().getResourceAsStream("test02.xlsx"))
                .columnSize(4)
                .sheetNum(IParserParam.FIRST_SHEET + 1)
                .targetClass(User.class)
                .header(User.getHeader())
                .build();

        List<User> user = parser.parse(parserParam);
        System.out.println(user);
    }

    @Test
    public void testCheckHeaderErrorXlsx() {
        parser = new ExcelSaxParser<>();

        IParserParam parserParam = DefaultParserParam.builder()
                .excelInputStream(Thread.currentThread().getContextClassLoader().getResourceAsStream("test02.xlsx"))
                .columnSize(4)
                .sheetNum(IParserParam.FIRST_SHEET + 1)
                .targetClass(User.class)
                .header(User.getErrHeader())
                .build();
        try{
            List<User> user = parser.parse(parserParam);
        }catch (Exception e){
            e.printStackTrace();
            return;
        }
        fail();
    }

    @Test
    public void testCheckHeaderErrorXls() {
        parser = new ExcelSaxParser<>();

        IParserParam parserParam = DefaultParserParam.builder()
                .excelInputStream(Thread.currentThread().getContextClassLoader().getResourceAsStream("test02.xls"))
                .columnSize(4)
                .sheetNum(IParserParam.FIRST_SHEET + 1)
                .targetClass(User.class)
                .header(User.getErrHeader())
                .build();
        try{
            List<User> user = parser.parse(parserParam);
        }catch (Exception e){
            e.printStackTrace();
            return;
        }
        fail();
    }

}
