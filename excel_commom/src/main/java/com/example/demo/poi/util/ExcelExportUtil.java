//
// Source code recreated from a .class file by IntelliJ IDEA
// (powered by Fernflower decompiler)
//

package com.example.demo.poi.util;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

public class ExcelExportUtil {
    public ExcelExportUtil() {
    }

    public byte[] exportExcel(Integer sheetnum, Map<Integer, List<? extends Object>> mapdata, Map<Integer, Map<String, Object>> mapheader) {
        XSSFWorkbook wb = new XSSFWorkbook();

        for(int i = 0; i <= sheetnum; ++i) {
            XSSFSheet sheet = wb.createSheet(Integer.toString(i));
            this.dealExcel(sheet, 0, (Map)mapheader.get(i), (List)mapdata.get(i));
        }

        ByteArrayOutputStream baos = null;

        Object var7;
        try {
            baos = new ByteArrayOutputStream();
            wb.write(baos);
            byte[] data = baos.toByteArray();
            byte[] var21 = data;
            return var21;
        } catch (Exception var17) {
            var17.printStackTrace();
            var7 = null;
        } finally {
            if (baos != null) {
                try {
                    baos.close();
                } catch (IOException var16) {
                    var16.printStackTrace();
                }
            }

        }

        return (byte[])var7;
    }

    //处理excel
    private void dealExcel(XSSFSheet sheet, Integer headerRowIndex, Map<String, Object> mapob, List<? extends Object> lobs, Object... objs) {
        Map<Integer, String> excelvalue_mapper = new HashMap();
        //初始行
        XSSFRow headerRow = sheet.createRow(headerRowIndex);
        //初始列
        Integer headerCellIndex = 0;
        //遍历标题的key
        Iterator var9 = mapob.keySet().iterator();

        while(var9.hasNext()) {
            String key = (String)var9.next();
            //得到的key,在初始行，初始列 ，添加标题
            headerRow.createCell(headerCellIndex).setCellValue(key);

            excelvalue_mapper.put(headerCellIndex, mapob.get(key).toString());
            if (headerCellIndex <= mapob.size()) {
                headerCellIndex = headerCellIndex + 1;
            }
        }

        var9 = lobs.iterator();

        while(var9.hasNext()) {
            Object ob = var9.next();
            XSSFRow contentRow = sheet.createRow(lobs.indexOf(ob) + headerRowIndex + 1);

            for(int cellIndex = 0; cellIndex <= headerCellIndex; ++cellIndex) {
                Field[] fields = ob.getClass().getDeclaredFields();
                Field[] var14 = fields;
                int var15 = fields.length;

                for(int var16 = 0; var16 < var15; ++var16) {
                    Field field = var14[var16];
                    field.setAccessible(true);
                    if (field.getName().equals(excelvalue_mapper.get(cellIndex))) {
                        try {
                            contentRow.createCell(cellIndex).setCellValue(String.valueOf(field.get(ob)));
                        } catch (Exception var19) {
                            var19.printStackTrace();
                        }
                    }
                }
            }
        }

    }

    public byte[] exportSimpleSheetExcel(Integer numSheet, Integer headerRowIndex, Map<String, Object> mapob, List<? extends Object> lobs, Object... objs) {
        Map<Integer, String> excelvalue_mapper = new HashMap();
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(numSheet.toString());
        XSSFRow headerRow = sheet.createRow(headerRowIndex);
        Integer headerCellIndex = 0;
        Iterator var11 = mapob.keySet().iterator();

        while(var11.hasNext()) {
            String key = (String)var11.next();
            headerRow.createCell(headerCellIndex).setCellValue(key);
            excelvalue_mapper.put(headerCellIndex, mapob.get(key).toString());
            if (headerCellIndex <= mapob.size()) {
                headerCellIndex = headerCellIndex + 1;
            }
        }

        var11 = lobs.iterator();

        XSSFRow contentRow;
        while(var11.hasNext()) {
            Object ob = var11.next();
            contentRow = sheet.createRow(lobs.indexOf(ob) + headerRowIndex + 1);

            for(int cellIndex = 0; cellIndex <= headerCellIndex; ++cellIndex) {
                Field[] fields = ob.getClass().getDeclaredFields();
                Field[] var16 = fields;
                int var17 = fields.length;

                for(int var18 = 0; var18 < var17; ++var18) {
                    Field field = var16[var18];
                    field.setAccessible(true);
                    if (field.getName().equals(excelvalue_mapper.get(cellIndex))) {
                        try {
                            contentRow.createCell(cellIndex).setCellValue(String.valueOf(field.get(ob)));
                        } catch (Exception var30) {
                            var30.printStackTrace();
                        }
                    }
                }
            }
        }

        ByteArrayOutputStream baos = null;

        try {
            baos = new ByteArrayOutputStream();
            wb.write(baos);
            byte[] data = baos.toByteArray();
            byte[] var36 = data;
            return var36;
        } catch (Exception var31) {
            var31.printStackTrace();
            contentRow = null;
        } finally {
            if (baos != null) {
                try {
                    baos.close();
                } catch (IOException var29) {
                    var29.printStackTrace();
                }
            }

        }

        return contentRow.toString().getBytes();
    }
}
