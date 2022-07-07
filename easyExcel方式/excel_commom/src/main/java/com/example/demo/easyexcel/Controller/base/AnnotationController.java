package com.example.demo.easyexcel.Controller.base;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.example.demo.easyexcel.model.Park;
import com.example.demo.easyexcel.model.User;
import com.example.demo.easyexcel.utils.DataUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;
import java.util.UUID;

/**
 * @program: excel_commom
 * @description: 注解式  适用于比较简单的 excel操作
 * @author: wjl
 * @create: 2022-07-03 15:43
 **/


/**
 * 说明：
 * 分页导出 每页行数上限是：65535
 * xls是03版的excel，行数限制6w+
 * xlsx是07版excel，行数限制100w+
 * 只需要更改后缀名，就可以导出两种不同格式的excel
 *
 * 如果需要点击下载，需要添加参数:HttpServletResponse response
 */
@RestController
public class AnnotationController {

    //注解式  分页excel  每页相同
    @RequestMapping("/annotationAndPagesSame")
    public String  annotationAndPages(){
        Integer pageNum = 5;//一共5页
        Integer num = 50;//每页50个
        //只需要更改后缀名，就可以导出两种不同格式的excel
        String path = "D:\\"+ UUID.randomUUID().toString().replace("-", "").toUpperCase()+".xlsx";
        ExcelWriter build = EasyExcel.write(path, User.class).build();
        //每50W条一个sheet页  xls不能达到50W条  注意调整
        for(int i = 0; i< pageNum ;i++){
            List<User> temp = DataUtils.getUsers(i,num);
            WriteSheet build1 = EasyExcel.writerSheet(i+"~"+(i+num)).build();
            build.write(temp,build1);
        }
        build.finish();
        System.out.println("导出成功！！！");
        return "导出成功！！！";
    }

    //注解式 分页excel 每页不同
    @RequestMapping("/annotationAndPagesDif")
    public String  annotationAndPagesDif(){
        //只需要更改后缀名，就可以导出两种不同格式的excel
        //String path = "D:\\"+UUID.randomUUID().toString().replace("-", "").toUpperCase()+".xls";
        String path = "D:\\"+UUID.randomUUID().toString().replace("-", "").toUpperCase()+".xlsx";
        List<User> temp = DataUtils.getUsers(0,50);
        List<Park> parks = DataUtils.getParks(0,100);
        ExcelWriter build = EasyExcel.write(path).build();
        WriteSheet sheet1 = EasyExcel.writerSheet("sheet1").head(User.class).build();
        WriteSheet sheet2 = EasyExcel.writerSheet("sheet2").head(Park.class).build();
        build.write(temp,sheet1);
        build.write(parks,sheet2);
        build.finish();
        System.out.println("导出成功！！！");
        return "导出成功！！！";
    }


}
