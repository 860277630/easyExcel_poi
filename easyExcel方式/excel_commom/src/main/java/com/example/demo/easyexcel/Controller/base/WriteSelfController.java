package com.example.demo.easyexcel.Controller.base;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.example.demo.easyexcel.utils.DataUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

/**
 * @program: excel_commom
 * @description: 非注解式 自定义开发 适用于比较复杂的excel操作
 * @author: wjl
 * @create: 2022-07-05 21:05
 **/
@RestController
public class WriteSelfController {

    //动态表头  就不用在实体前加注解了
    @RequestMapping("selfTitleSame")
    public String exportBySelfTitle(){

        //1.用实体的方式
        // List<Student> list = getList();
        //2.用非实体方式
        List<List<String>> list = DataUtils.getResolveList();

        List<List<String>> heads = new ArrayList<>();
        List<String> head1 = new ArrayList<>();
        head1.add("基本信息");
        head1.add("ID");
        List<String> head2 = new ArrayList<>();
        head2.add("基本信息");
        head2.add("姓名");
        List<String> head3 = new ArrayList<>();
        head3.add("基本信息");
        head3.add("性别");
        List<String> head4 = new ArrayList<>();
        head4.add("教育信息");
        head4.add("年级");
        List<String> head5 = new ArrayList<>();
        head5.add("教育信息");
        head5.add("老师");

        heads.add(head1);
        heads.add(head2);
        heads.add(head3);
        heads.add(head4);
        heads.add(head5);

        String fileName = "D:\\"+ UUID.randomUUID().toString().replace("-", "").toUpperCase()+".xlsx";

        //分批导出
        ExcelWriter build = EasyExcel.write(fileName).head(heads).build();
        WriteSheet sheet = EasyExcel.writerSheet("学生信息").build();
        for(int i = 0; i< list.size();i+=500){
            List<List<String>> temp = list.subList(i,(i+500)<list.size()?(i+500):list.size());
            build.write(temp,sheet);
        }
        build.finish();
        //整个一起导出
        //ExcelWriterBuilder builder = EasyExcel.write(fileName);
        //builder.head(heads).sheet("学生信息").doWrite(list);
        System.out.println("导出成功！！！");
        return "export";
    }


    //不一样的sheet 导出不一样的数据 并且自写 不采用实体填充
    @RequestMapping("selfTitleDiff")
    public String exportByDifferenceSheetAndSelf(){
        String fileName = "D:\\"+UUID.randomUUID().toString().replace("-", "").toUpperCase()+".xlsx";

        //1.用实体的方式
        // List<Student> list = getList();
        //2.用非实体方式
        List<List<String>> list = DataUtils.getResolveList();
        List<List<String>> heads = new ArrayList<>();
        List<String> head1 = new ArrayList<>();
        head1.add("基本信息");
        head1.add("ID");
        List<String> head2 = new ArrayList<>();
        head2.add("基本信息");
        head2.add("姓名");
        List<String> head3 = new ArrayList<>();
        head3.add("基本信息");
        head3.add("性别");
        List<String> head4 = new ArrayList<>();
        head4.add("教育信息");
        head4.add("年级");
        List<String> head5 = new ArrayList<>();
        head5.add("教育信息");
        head5.add("老师");

        heads.add(head1);
        heads.add(head2);
        heads.add(head3);
        heads.add(head4);
        heads.add(head5);
        //分批导出
        ExcelWriter build = EasyExcel.write(fileName).build();
        for(int i = 0; i< list.size();i+=500){
            WriteSheet sheet = EasyExcel.writerSheet("学生信息"+i).head(heads).build();
            List<List<String>> temp = list.subList(i,(i+500)<list.size()?(i+500):list.size());
            build.write(temp,sheet);
        }

        //然后再加一个特别的sheet
        List<Object> data = new ArrayList<>();
        List<List<String>> header = DataUtils.head(data);
        List<List<Object>> lists = new ArrayList<>();
        lists.add(data);
        WriteSheet sheet = EasyExcel.writerSheet("special").head(header).build();
        build.write(lists,sheet);
        //完成
        build.finish();
        System.out.println("导出成功！！！");
        return "export";
    }


}
