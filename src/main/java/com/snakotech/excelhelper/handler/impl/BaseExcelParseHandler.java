package com.snakotech.excelhelper.handler.impl;

import com.snakotech.excelhelper.IParserParam;
import com.snakotech.excelhelper.handler.IExcelParseHandler;
import com.snakotech.excelhelper.meta.ExcelField;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

abstract class BaseExcelParseHandler<T> implements IExcelParseHandler<T> {

    Logger logger = LoggerFactory.getLogger(this.getClass());
    private boolean head = true;

    <T> Optional<T> parseRowToTarget(IParserParam parserParam, List<String> rowData) {
        if (isRowDataEmpty(rowData)) {
            return Optional.empty();
        }

        try {
            T t = doParseRowToTarget(rowData, parserParam.getTargetClass());
            return Optional.of(t);
        } catch (Exception e) {
            logger.error("AbstractExcelParser - parseRowToTarget", e);
            return Optional.empty();
        }
    }


    private <T> T doParseRowToTarget(List<String> rowData, Class targetClass) throws IllegalAccessException, InstantiationException {
        T object = (T) targetClass.newInstance();
        Field[] declaredFields = object.getClass().getDeclaredFields();
        for (Field field : declaredFields) {
            if (field.isAnnotationPresent(ExcelField.class)) {
                ExcelField excelField = field.getAnnotation(ExcelField.class);
                int index = excelField.index();
                ExcelField.ExcelFieldType type = excelField.type();
                field.setAccessible(true);
                String setValue = rowData.get(index);
                String toSet = type.buildSetString(setValue);
                field.set(object, toSet);
            }
        }
        return object;
    }

    List<String> initRowList(int size) {
        List<String> list = new ArrayList<>();
        for (int i = 0; i < size; i++)
            list.add("");
        return list;
    }

    boolean isRowDataEmpty(List<String> rowData) {
        if (rowData == null)
            return true;
        for (String str : rowData) {
            if (str != null && !Objects.equals("", str.trim())) {
                return false;
            }
        }
        return true;
    }

    void validHeader(IParserParam parserParam, List<String> rowData) {
        int index = 0;
        if (rowData.size() != parserParam.getHeader().size()) {
            throw new IllegalArgumentException("Excel Header Check Failed");
        }
        for (String head : parserParam.getHeader()) {
            if (!Objects.equals(rowData.get(index++), head.trim())) {
                throw new IllegalArgumentException("Excel Header Check Failed");
            }
        }
    }

    protected void handleEndOfRow(IParserParam parserParam, List<String> rowData, List<T> result) {
        boolean empty = isRowDataEmpty(rowData);
        if (!empty) {
            if (head && parserParam.getHeader() != null && parserParam.getHeader().size() != 0) {
                validHeader(parserParam, rowData);
                head = false;
            } else {
                Optional<T> t = parseRowToTarget(parserParam, rowData);
                t.ifPresent(result::add);
            }
        }
    }
}
