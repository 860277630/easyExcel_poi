package com.example.demo.poi.util;

import cn.binarywang.tools.generator.ChineseNameGenerator;
import com.example.demo.poi.model.User;

import java.util.*;

/**
 * @program: excel_commom
 * @description: 数据提供工具类
 * @author: wjl
 * @create: 2022-07-03 16:34
 **/
public class DataUtils {

    public static List<User> getUsers(Integer page,Integer num){
        List<User> result = new ArrayList<>();
        for (Integer order = 0; order < num; order++) {
            result.add(new User((page-1)*num+order,getName(), UUID.randomUUID().toString(),getName()));
        }
        return result;
    }



    //生成随机姓名
    private static String getName(){
        ChineseNameGenerator instance = ChineseNameGenerator.getInstance();
        return  instance.generate();
    }





}
