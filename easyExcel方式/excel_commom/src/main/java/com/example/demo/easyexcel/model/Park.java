package com.example.demo.easyexcel.model;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;

/**
 * park
 * @author 
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class Park {


    @ColumnWidth(8)
    @ExcelProperty(index = 0,value = "公园ID")
    private String id;


    @ColumnWidth(8)
    @ExcelProperty(index = 1,value = "公园名字")
    private String name;


    @ColumnWidth(8)
    @ExcelProperty(index = 2,value = "城市名字")
    private String city;


    @ColumnWidth(8)
    @ExcelProperty(index = 3,value = "面积")
    private String area;

    @ExcelIgnore
    private String miaoshu;

    /**
     * 公园的树
     */
    @ExcelProperty(index = 4,value = "树名")
    private String trees;

}