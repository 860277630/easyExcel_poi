package com.example.demo.easyexcel.mergeCls;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import lombok.Data;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * @program: easyExcel
 * @description: 自定义excel合并类
 * @author: wjl
 * @create: 2022-06-20 17:33
 **/

@Data
public class ExcelFillCellMergeBySelf implements CellWriteHandler {

    //假定：原先的excel 并没有合并区域 都是行列清晰的excel
    //不考虑复杂情况  如：合并区域内原先就有合并区域  或者合并区域分割了原先的合并区域

    //模式  分两个模式
    //0 :统计区域中所有的字符串 ，放于合并后的单元格内
    //1 :对于合并区域除第左上第一个单元格外，其他的单元格向右或者向下进行推移
    private int mergeModel ;

    //合并区域
    //四个坐标分别是：从上到下，从左到右，起始列，终止列，起始行，终止行
    private int[] mergeArexa;

    //用来统计  excel 文本的容器
    private static List<String> cellCens = new ArrayList<>();

    public ExcelFillCellMergeBySelf(){}

    /**
     *
     * @param mergeModel
     * @param mergeArexa  四个坐标分别是：从上到下，从左到右，起始列，终止列，起始行，终止行
     */
    public ExcelFillCellMergeBySelf(int mergeModel,int[] mergeArexa){
        this.mergeModel = mergeModel;
        this.mergeArexa = mergeArexa;
    }


    //在创建单元格之前调用
    @Override
    public void beforeCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Head head, Integer columnIndex, Integer relativeRowIndex, Boolean isHead) {
        //System.out.println("******come in beforeCellCreate******");
    }

    //创建单元格后调用
    @Override
    public void afterCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        // System.out.println("******come in afterCellCreate******");
    }

    //在单元格上的所有操作完成后调用
    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<WriteCellData<?>> list,
                                 Cell cell, Head head, Integer integer, Boolean aBoolean) {
        //System.out.println("***come in afterCellDispose***");
        //进行遍历 合并

        //判断当前行和列是否在  合并区域内  如果再合并区域内
        //如果在合并区域内 就根据模式
        //0模式：收集格区域内所有文字 进行统一展示
        //1模式：将区域内的文字进行向右移动
        int rowIndex = cell.getRowIndex();
        int columnIndex = cell.getColumnIndex();
        //如果在区间内根据模式进行操
        if(columnIndex>=mergeArexa[0]&&columnIndex<=mergeArexa[1]
        &&rowIndex>=mergeArexa[2]&&rowIndex<=mergeArexa[3]){
            if(mergeModel == 0){
                //收集所有的单元格文本 直到合并后再填入
                //最后一个单元格 行为mergeArexa[1]  列为 行为mergeArexa[3]
                cellCens.add(String.valueOf(cell.getCellTypeEnum() == CellType.STRING ? cell.getStringCellValue() : cell.getNumericCellValue()));
                //到了最后一个单元格 就把合并区域加入进去 并且 进行文本的填充
                if(rowIndex == mergeArexa[1] && columnIndex == mergeArexa[3]){
                    //因为easyexcel的合并都是取左上
                    //所以需要将所有的数据 堆到 左上的单元格内部 在进行合并
                    cell.getSheet().getRow(mergeArexa[2]).getCell(mergeArexa[0]).setCellValue(StringUtils.join(cellCens,""));
                    //清空内存
                    cellCens.clear();
                    //把合并区域加入到 合并列表中 并且为 合并区域设置新值
                    Sheet sheet = writeSheetHolder.getSheet();
                    CellRangeAddress cellRangeAddress = new CellRangeAddress(mergeArexa[2], mergeArexa[3], mergeArexa[0], mergeArexa[1]);
                    sheet.addMergedRegion(cellRangeAddress);
                }
            }
            else if(mergeModel == 1){
                //模式1：将合并的 区域置为空  其他地方向后推移
                //只要符合条件就进行数据的收集
                cellCens.add(String.valueOf(cell.getCellTypeEnum() == CellType.STRING ? cell.getStringCellValue() : cell.getNumericCellValue()));
                //符合最后一个单元格 就进行下列处理
                //1.合并区域 全部向后移动   并且合并区域内置为空
                //2.添加合并区域
                if(rowIndex == mergeArexa[1] && columnIndex == mergeArexa[3]){
                    for(Integer startRow = mergeArexa[2];startRow<=mergeArexa[3];startRow++){
                        //获取列 然后 迭代向后推进
                        Row row = cell.getSheet().getRow(startRow);
                        //先用迭代器拿出来 然后再用for循环 进行row的修改
                        List<Cell> tempList = new ArrayList<>();
                        Iterator<Cell> iterator = row.cellIterator();
                        while (iterator.hasNext()){
                            tempList.add(iterator.next());
                        }
                        //得到了全部数据后 进行row的填充
                        Integer value = Integer.valueOf(row.getLastCellNum());
                        int width = mergeArexa[1] - mergeArexa[0];

                        for(Integer startCol = mergeArexa[0];startCol <=(value+width);startCol++){
                            if((startCol - mergeArexa[0])<=width){
                                Cell tempCell = row.createCell(startCol);
                                tempCell.setCellValue(StringUtils.EMPTY);

                            }else {
                                //拿到对应的值
                                Cell cel = tempList.get(startCol - width-1);
                                Cell rowCell = row.createCell(startCol);
                                rowCell.setCellValue(cel.getStringCellValue());
                            }
                        }
                    }
                    //这样的话 就得到了 推移后的excel 然后进行合并就可以了
                    Sheet sheet = writeSheetHolder.getSheet();
                    CellRangeAddress cellRangeAddress = new CellRangeAddress(mergeArexa[2], mergeArexa[3], mergeArexa[0], mergeArexa[1]);
                    sheet.addMergedRegion(cellRangeAddress);
                }
            }
            else{
                System.out.println("rowIndex:"+rowIndex+"columnIndex:"+columnIndex);
                cell.setCellValue("王龙");
            }
        }
    }
}
