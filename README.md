# SnokeTech Excel Parser

## 说明
这是一个轻量级的Excel解析框架.

## 使用
```java
        IExcelParser<User> parser = new ExcelDomParser<>();

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

```

## 联系方式

QQ:398110112