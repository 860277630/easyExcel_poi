package com.example.demo.easyexcel.Controller.download;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.example.demo.easyexcel.mergeCls.ExcelFillCellMergeStrategy;
import com.example.demo.easyexcel.model.DemoData;
import com.example.demo.easyexcel.utils.DataUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.BufferedOutputStream;
import java.net.URLEncoder;
import java.util.List;

/**
 * @program: excel_commom
 * @description: excel下载到自定义位置
 * @author: wjl
 * @create: 2022-07-07 20:45
 **/

//如果要下载到自定义的位置  加载到HttpServletResponse  即可
@RestController
public class DownLoadController {

    @RequestMapping("downloadBySelf")
    public String mergeAdjacentAndIdentical(HttpServletResponse response) throws Exception{
        //获取数据源
        List<DemoData> data = DataUtils.data();
        //设置输入流，设置响应域
        response.setContentType("application/ms-excel");
        response.setCharacterEncoding("utf-8");
        String  fileName = URLEncoder.encode("ceshi.xlsx","utf-8");
        response.setHeader("Content-disposition","attachment;filename="+fileName);
        //需要合并的列
        int[] mergeColumeIndex = {0, 1, 2};
        //需要从第一行开始，列头第一行
        int mergeRowIndex = 1;

        EasyExcel//将数据映射到DownloadDTO实体类并响应到浏览器
                .write(new BufferedOutputStream(response.getOutputStream()), DemoData.class)
                //07的excel版本,节省内存
                .excelType(ExcelTypeEnum.XLSX)
                //是否自动关闭输入流
                .autoCloseStream(Boolean.TRUE)
                .registerWriteHandler(new ExcelFillCellMergeStrategy(mergeRowIndex, mergeColumeIndex))
//               // 自定义列宽度，有数字会
//                .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())
                //设置excel保护密码
//                .password("123456")
                .sheet().doWrite(data);

        System.out.println("导出成功！！！");
        return "export";
    }

}
