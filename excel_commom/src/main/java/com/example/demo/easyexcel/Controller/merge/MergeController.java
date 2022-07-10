package com.example.demo.easyexcel.Controller.merge;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.merge.LoopMergeStrategy;
import com.alibaba.excel.write.merge.OnceAbsoluteMergeStrategy;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.example.demo.easyexcel.mergeCls.ExcelFillCellMergeBySelf;
import com.example.demo.easyexcel.mergeCls.ExcelFillCellMergeStrategy;
import com.example.demo.easyexcel.utils.DataUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

/**
 * @program: excel_commom
 * @description: 用于各种表格的合并，当然对于过于复杂的excel还是建议用模板标记
 *
 * 固定合并方法
 * 使用OnceAbsoluteMergeStrategy  以及  LoopMergeStrategy  对excel进行合并  合并后单元格的值 同左上 第一个单元格的值一致
 * OnceAbsoluteMergeStrategy 单一区域合并  构造函数 firstRowIndex: 起始行 lastRowIndex：终止行 firstColumnIndex：起始列 lastColumnIndex：终止列
 * LoopMergeStrategy  迭代式合并 构造函数  eachRow：每几行合并成一部分 columnExtend：合并范围，跨N列 columnIndex：合并开始列
 *
 * 自定义合并方法
 * 重写 CellWriteHandler 和 RowWriteHandler  、SheetWriteHandler 进行合并
 * 满足更多合并要求 可以继承接口进行重写    cell 单元格   Row 行   Sheet 页
 *
 * @author: wjl
 * @create: 2022-07-05 22:16
 **/

//对于复杂表格 还是建议使用easyexcel模板 此处仅仅是对easy excel复杂使用的展示
@RestController
public class MergeController {
    @RequestMapping("divMerge")
    public String exportBymerge() throws IOException {
        //获取数据源
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
        //需要合并的列
        int[] mergeColumeIndex = {0, 1, 2};
        //需要从第一行开始，列头第一行
        int mergeRowIndex = 1;

        //指定 特定的行和列进行合并
        //从第2列开始 ，每3行进行一次合并，合并范围为3
        //合并的单元格值 与首个单元格值一致
        LoopMergeStrategy loopMergeStrategy = new LoopMergeStrategy(3,3,2);

        //指定特定的行和列进行合并
        //单一区域，非迭代
        //从第5行到第8行 第2列 到 第5列 进行合并
        //合并的单元格值 与首个单元格值一致
        OnceAbsoluteMergeStrategy onceAbsoluteMergeStrategy = new OnceAbsoluteMergeStrategy(5,8,2,5);


        String fileName = "D:\\"+ UUID.randomUUID().toString().replace("-", "").toUpperCase()+".xlsx";

        //导出到两个不同的sheet页
        ExcelWriter build = EasyExcel.write(fileName)
                .excelType(ExcelTypeEnum.XLSX).autoCloseStream(Boolean.TRUE).build();

        //创建sheet页
        WriteSheet sheet_1 = EasyExcel.writerSheet("完整信息").head(heads).build();

        //这种合并是 同列上相同值进行合并
        WriteSheet sheet_2 = EasyExcel.writerSheet("相同值进行合并").head(heads)
                .registerWriteHandler(new ExcelFillCellMergeStrategy(mergeRowIndex, mergeColumeIndex)).build();

        //自定义合并1
        WriteSheet sheet_3 = EasyExcel.writerSheet("自定义合并1_迭代合并").head(heads)
                .registerWriteHandler(loopMergeStrategy).build();


        //自定义合并2
        WriteSheet sheet_4 = EasyExcel.writerSheet("自定义合并2_单一合并").head(heads)
                .registerWriteHandler(onceAbsoluteMergeStrategy).build();


        //自定义合并3
        //指定合并的行和列 ，由于提供的OnceAbsoluteMergeStrategy 合并后的值取的是左上第一个单元格的值
        //这里进行特别处理 对于合并区域被覆盖的值
        //网上没有找到简便的办法， 这里用过滤器来做
        //1.第一个合并区域，统计区域中所有的字符串 ，放于合并后的单元格内
        //2.第二个合并区域，对于合并区域除第左上第一个单元格外，其他的单元格向右或者向下进行推移
        WriteSheet sheet_5 = EasyExcel.writerSheet("自定义合并3_自定义合并").head(heads)
                .registerWriteHandler(new ExcelFillCellMergeBySelf(1, new int[]{2, 4, 2, 4})).build();

        //进行组装
        build.write(list,sheet_1);
        build.write(list,sheet_2);
        build.write(list,sheet_3);
        build.write(list,sheet_4);
        build.write(list,sheet_5);

        //关闭流
        build.finish();
        System.out.println("导出成功！！！");
        return "export";
    }
}
