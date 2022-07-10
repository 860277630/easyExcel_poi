package com.example.demo.poi.Controller;

import com.example.demo.poi.util.DataUtils;
import com.example.demo.poi.model.Books;
import com.example.demo.poi.model.User;
import com.example.demo.poi.util.ExcelExportUtil;
import com.example.demo.poi.util.ExcelTransUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * @program: excel_commom
 * @description: poi操作excel
 * @author: wjl
 * @create: 2022-07-10 16:22
 **/

@RestController
@Slf4j
public class PoiController {


    @RequestMapping("readXls")
    public String readXls(){
        //获取当前路径
        String filePath = System.getProperty("user.dir")+"\\src\\main\\resources\\static\\Excels\\books.xls";
        List<Books> list = new ArrayList<Books>();
        try {
            FileInputStream fis = new FileInputStream(filePath);
            POIFSFileSystem fs = new POIFSFileSystem(fis);
            HSSFWorkbook workbook = new HSSFWorkbook(fs);
            // 遍历excel中的所有表
            for (int sheetnum = 0; sheetnum < workbook.getNumberOfSheets(); sheetnum++) {
                HSSFSheet sheet = workbook.getSheetAt(sheetnum);
                for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                    Books books = new Books();
                    HSSFRow row = sheet.getRow(i);
                    HSSFCell cell1 = row.getCell(0);
                    cell1.setCellType(CellType.STRING);// 将格式设置为string，不然如果excel存的是数字是取不出来的

                    System.out.println(cell1.getStringCellValue());

                    HSSFCell cell2 = row.getCell(1);
                    cell2.setCellType(CellType.STRING);// 将格式设置为string，不然如果excel存的是数字是取不出来的

                    System.out.println(cell2.getStringCellValue());

                    HSSFCell cell3 = row.getCell(2);
                    cell3.setCellType(CellType.STRING);// 将格式设置为string，不然如果excel存的是数字是取不出来的
                    System.out.println(cell3.getStringCellValue());

                    HSSFCell cell4 = row.getCell(3);
                    cell4.setCellType(CellType.STRING);// 将格式设置为string，不然如果excel存的是数字是取不出来的
                    System.out.println(cell4.getStringCellValue());


                    HSSFCell cell5 = row.getCell(4);
                    cell5.setCellType(CellType.STRING);// 将格式设置为string，不然如果excel存的是数字是取不出来的
                    System.out.println(cell5.getStringCellValue());

                    books.setId(Integer.parseInt(cell1.getStringCellValue()));
                    books.setBookname(cell2.getStringCellValue());
                    books.setPrices(Double.parseDouble(cell3.getStringCellValue()));
                    books.setCounts(Integer.parseInt(cell4.getStringCellValue()));
                    books.setTypeid(Integer.parseInt(cell5.getStringCellValue()));
                    list.add(books);
                }
            }
            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }

        //最后遍历输出
        log.info("每行的数据如下：");
        list.forEach(x->{
            System.out.println(x.toString());
        });
        return "处理完毕";
    }

    @RequestMapping("readXlsx")
    public String readxlsx(){
        List<Books> list = new ArrayList<Books>();
        String filePath = System.getProperty("user.dir")+"\\src\\main\\resources\\static\\Excels\\books.xlsx";

        try {
            FileInputStream fis = new FileInputStream(filePath);
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            System.out.println(workbook.getNumberOfSheets());
            // 遍历excel中的所有表
            for (int sheetnum = 0; sheetnum < workbook.getNumberOfSheets(); sheetnum++) {

                XSSFSheet sheet = workbook.getSheetAt(0);
                for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                    Books books = new Books();
                    XSSFRow row = sheet.getRow(i);
                    XSSFCell cell1 = row.getCell(0);
                    cell1.setCellType(CellType.STRING);// 将格式设置为string，不然如果excel存的是数字是取不出来的
                    XSSFCell cell2 = row.getCell(1);
                    cell2.setCellType(CellType.STRING);// 将格式设置为string，不然如果excel存的是数字是取不出来的
                    XSSFCell cell3 = row.getCell(2);
                    cell3.setCellType(CellType.STRING);// 将格式设置为string，不然如果excel存的是数字是取不出来的
                    XSSFCell cell4 = row.getCell(3);
                    cell4.setCellType(CellType.STRING);// 将格式设置为string，不然如果excel存的是数字是取不出来的
                    XSSFCell cell5 = row.getCell(4);
                    cell5.setCellType(CellType.STRING);// 将格式设置为string，不然如果excel存的是数字是取不出来的

                    books.setId(Integer.parseInt(cell1.getStringCellValue()));
                    books.setBookname(cell2.getStringCellValue());
                    books.setPrices(Double.parseDouble(cell3.getStringCellValue()));
                    books.setCounts(Integer.parseInt(cell4.getStringCellValue()));
                    books.setTypeid(Integer.parseInt(cell5.getStringCellValue()));
                    list.add(books);
                }
            }
            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
        //最后遍历输出
        log.info("每行的数据如下：");
        list.forEach(x->{
            System.out.println(x.toString());
        });
        return "处理完毕";
    }

    //导出excel  自定义下载位置
    @RequestMapping("exportExcelBySelfPath")
    public ResponseEntity<byte[]> exportExcelBySelfPath(){
        List<User> users = DataUtils.getUsers(1,10);

        Map<String, Object> map = new LinkedHashMap();
        //这是列名及实体名
        map.put("用户ID", "id");
        map.put("用户名", "userName");
        map.put("密码", "passWord");
        map.put("用户权限", "realName");
        ExcelExportUtil ee = new ExcelExportUtil();
        byte[] exceldata = ee.exportSimpleSheetExcel(0, 0, map, users, new Object[0]);

        /*
         *参数说明：
         *1.导入数据包
         *
         * 2.文件名
         * */
        return ExcelTransUtil.exportFile(exceldata, "用户权限表");
    }

}
