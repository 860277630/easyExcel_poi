package com.example.demo.easyexcel.model;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class User {
    @ExcelProperty(index = 0,value = {"主标题", "用户ID"})
    private Integer id;
    @ExcelProperty(index = 2,value = {"主标题", "用户名"})
    private String userName;
    @ExcelProperty(index = 1,value = {"主标题", "密码"})
    private String passWord;
    //不用输出
    @ExcelIgnore
    private String realName;
}