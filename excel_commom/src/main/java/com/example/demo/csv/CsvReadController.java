package com.example.demo.csv;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * @program: excel_commom
 * @description: 读取csv
 * @author: wjl
 * @create: 2023-04-25 09:35
 **/
@RestController
public class CsvReadController {

    @RequestMapping("readCsv")
    public String readCsv(){
        String filePath = System.getProperty("user.dir")+"\\src\\main\\resources\\static\\Excels\\res_end.csv";
        String line = "";
        String SplitBy = ",";
        String[] Line;
        List<String[]> BikeDataList = new ArrayList<>();
        try (BufferedReader br = new BufferedReader(new FileReader(filePath))) {
            while ((line = br.readLine()) != null) {
                Line = line.split(SplitBy);
                BikeDataList.add(Line);
                System.out.println("signal "+Line[0]+" {\n\tsite 1 { pogo = "+Line[1]+"; }" +
                        "\n\tsite 2 { pogo = "+Line[2]+"; }"+"\n\tsite 3 { pogo = "+Line[3]+"; }"+"\n}");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return "sucess";
    }

}
