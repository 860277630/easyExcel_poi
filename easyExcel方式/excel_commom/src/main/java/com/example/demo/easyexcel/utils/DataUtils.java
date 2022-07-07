package com.example.demo.easyexcel.utils;

import cn.binarywang.tools.generator.ChineseAddressGenerator;
import cn.binarywang.tools.generator.ChineseAreaList;
import cn.binarywang.tools.generator.ChineseNameGenerator;
import com.example.demo.easyexcel.model.DemoData;
import com.example.demo.easyexcel.model.Park;
import com.example.demo.easyexcel.model.Student;
import com.example.demo.easyexcel.model.User;

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

    public static List<Park> getParks(Integer page,Integer num){
        List<Park> result = new ArrayList<>();
        for (Integer order = 0; order < num; order++) {
            result.add(new Park(UUID.randomUUID().toString()
                    ,getParkName()
                    ,getCity()
                    ,getArea()
                    ,getAdress()
                    ,"芭蕉树"));
        }
        return result;
    }

    public static Map<String, List<String>> getHeader() {
        Map<String, List<String>> map = new HashMap<>();
        List<String> aList = new ArrayList<>();
        List<String> sList = new ArrayList<>();
        List<String> subList = new ArrayList<>();

        for (int i = 0; i < 10; i++) {
            String column = "温度" + i;
            aList.add(column);
            String scolumn = "风门" + i;
            sList.add(scolumn);
        }
        aList.add("小计1");
        sList.add("小计2");
        String subColumn = "其它";
        subList.add(subColumn);
        subList.add("小计3");
        map.put("模拟量传感器", aList);
        map.put("开关量传感器", sList);
        return map;
    }

    //获取数据的地方
    public static List<DemoData> data() {
        List<DemoData> list = new ArrayList<DemoData>();
        for (int i = 0; i < 10; i++) {
            DemoData data = new DemoData();
            data.setString("字符串" + i);
            data.setDate(new Date());
            data.setDoubleData(0.56);
            list.add(data);
        }
        return list;
    }


    /**
     * 返回的数据结构如下
     * {"data":{"xxxxxxx":{"":32,"1011":42,"1010":8,"1008":52,"0006":3,"0004":86,"1004":2,"0013":7,"0002":7,"1003":4,"0003":56,"1002":18,"0011":13,"1001":12,"0012":6,"0001":66,"1009":42}},
     * "deviceTypeSet":{"switching":["1008","1004","1003","1002","1001","1011","1010"],
     * "analog":["0006","0004","0013","0002","0003","0011","0012","0001"],
     * "switchingOff":["1009"],"substation":[""]},
     * "success":true}
     *
     * @return
     */
    //解析表头  同时填充data  表头和data是互不相干的
    public static List<List<String>> head(List<Object> data) {
        List<List<String>> list = new ArrayList<List<String>>();
        List<String> head0 = new ArrayList<String>();
        head0.add("机构名称");
        list.add(head0);
        data.add("张璐的矿");
        Map<String, List<String>> map = getHeader();
        Map<String, Long> dataMap = getData();
        map.forEach((k, v) -> {
            String deviceCategory = k;
            List<String> ls = v;
            ls.forEach(e -> {
                List<String> head = new ArrayList<>();
                head.add(deviceCategory);
                head.add(e);
                list.add(head);
                if (dataMap.containsKey(e)) {
                    data.add(dataMap.get(e));
                } else {
                    data.add(0);
                }
            });

        });
        List<String> headn = new ArrayList<String>();
        headn.add("合计");
        list.add(headn);
        data.add(dataMap.get("合计"));
        return list;
    }


    private static Map<String, Long> getData() {
        //{"10017904-1":
        //{"":32,"1011":42,"1010":8,"1008":52,"0006":3,"0004":86,"1004":2,"0013":7,"0002":7,"1003":4,"0003":56,"1002":18,"0011":13,"1001":12,"0012":6,"0001":66,"1009":42}},
        Map<String, Long> data = new HashMap<>();
        long atotal = 0;
        long stotal = 0;
        long subtotal = 0;
        for (int i = 0; i < 10; i++) {
            String column = "温度" + i;
            atotal += Math.round(Math.random()) + i;
            data.put(column, Math.round(Math.random()) + i);
            String scolumn = "风门" + i;
            stotal += Math.round(Math.random()) + i;
            data.put(scolumn, Math.round(Math.random()) + i);
        }
        data.put("小计1", atotal);
        data.put("小计2", stotal);
        String subColumn = "其它";
        data.put(subColumn, 55l);
        data.put("小计3", 55l);
        data.put("合计", 55 + atotal + stotal);
        return data;
    }

    private static List<Student>  getList(){
        List<Student> list = new ArrayList<>();
        for(int i = 0; i< 1000; i++){
            list.add(new Student(String.valueOf(i),getName(),getSex(),getGrade(),getName()));
        }
        return list;
    }


    public static List<List<String>> getResolveList(){
        List<List<String>> result = new ArrayList<>();
        List<Student> list = getList();
        //对其进行分解
        list.forEach(x->{
            List<String> temp = new ArrayList<>();
            temp.add(x.getId());
            temp.add(x.getName());
            temp.add(x.getSex());
            temp.add(x.getGrade());
            temp.add(x.getTeacher());
            result.add(temp);
        });
        return result;
    }



    //生成随机姓名
    private static String getName(){
        ChineseNameGenerator instance = ChineseNameGenerator.getInstance();
        return  instance.generate();
    }
    //生成随机性别
    private static String getSex() {
        int randNum = new Random().nextInt(2) + 1;
        return randNum == 1 ? "男" : "女";
    }
    //生成随机年级
    private static String getGrade(){
        return String.valueOf(new Random().nextInt(6) + 1) ;
    }

    //生成随机地址
    private static String getAdress(){
        return ChineseAddressGenerator.getInstance().generate();
    }

    //生成随机城市
    private static String getCity(){
        List<String> cityList = ChineseAreaList.provinceCityList;
        int randNum = new Random().nextInt(cityList.size()) + 1;
        return ChineseAreaList.provinceCityList.get(randNum);
    }
    //生成随机公园名称
    private static String getParkName(){
        String generate = ChineseAddressGenerator.getInstance().generate();
        return generate.substring(generate.indexOf("号")+1, generate.indexOf("小区"))+"公园";
    }

    //生成随机面积
    private static String getArea(){
        int randNum = new Random().nextInt(20000) + 1;
        return String.valueOf(randNum)+"平方米";
    }

    public static void main(String[] args) {
        System.out.println(getCity());
        System.out.println(getName());
        System.out.println(getParkName());
    }


}
