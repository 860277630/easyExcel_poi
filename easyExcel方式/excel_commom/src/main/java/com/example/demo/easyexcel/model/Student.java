package com.example.demo.easyexcel.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;

/**
 * student
 * @author 
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class Student {
    private String id;

    private String name;

    private String sex;

    private String grade;

    private String teacher;
}