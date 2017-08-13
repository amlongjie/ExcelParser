package com.snakotech;

import com.snakotech.excelhelper.meta.ExcelField;

import java.util.ArrayList;
import java.util.List;

public class User {

    @ExcelField(index = 0)
    private String name;
    @ExcelField(index = 1)
    private String age;
    @ExcelField(index = 2)
    private String gender;
    @ExcelField(index = 3, type = ExcelField.ExcelFieldType.Date)
    private String dateStr;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getAge() {
        return age;
    }

    public void setAge(String age) {
        this.age = age;
    }

    public String getGender() {
        return gender;
    }

    public void setGender(String gender) {
        this.gender = gender;
    }

    public String getDateStr() {
        return dateStr;
    }

    public void setDateStr(String dateStr) {
        this.dateStr = dateStr;
    }

    @Override
    public String toString() {
        return "User{" +
                "name='" + name + '\'' +
                ", age='" + age + '\'' +
                ", gender='" + gender + '\'' +
                ", dateStr='" + dateStr + '\'' +
                '}';
    }

    public static List<String> getHeader() {
        List<String> header = new ArrayList<>();
        header.add("姓名");
        header.add("年龄");
        header.add("性别");
        header.add("生日");
        return header;
    }

    public static List<String> getErrHeader() {
        List<String> header = new ArrayList<>();
        header.add("姓名");
        return header;
    }
}
