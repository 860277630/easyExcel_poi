//
// Source code recreated from a .class file by IntelliJ IDEA
// (powered by Fernflower decompiler)
//

package com.example.demo.poi.util;

import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;

import javax.servlet.http.HttpServletRequest;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;

public class ExcelTransUtil {
    public ExcelTransUtil() {
    }

    public static ResponseEntity<byte[]> exportFile(byte[] data, String filename) {
        try {
            filename = new String(filename.getBytes("GBK"), "iso-8859-1");
        } catch (UnsupportedEncodingException var3) {
            var3.printStackTrace();
        }

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        headers.setContentDispositionFormData("attachment", filename + ".xlsx");
        return new ResponseEntity(data, headers, HttpStatus.CREATED);
    }

    public static ResponseEntity<byte[]> exportFile2(byte[] data, String filename, HttpServletRequest request) {
        try {
            String userAgent = request.getHeader("User-Agent");
            if (!userAgent.contains("MSIE") && !userAgent.contains("Trident")) {
                filename = new String(filename.getBytes("UTF-8"), "ISO-8859-1");
            } else {
                filename = URLEncoder.encode(filename, "UTF-8");
            }
        } catch (UnsupportedEncodingException var4) {
            var4.printStackTrace();
        }

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        headers.setContentDispositionFormData("attachment", filename + ".xlsx");
        return new ResponseEntity(data, headers, HttpStatus.CREATED);
    }

    /*public static byte[] exportDataExcel(Integer numSheet, Integer headerRowIndex, List<Map<String, String>> datas) {
        if (CollectionUtils.isEmpty(datas)) {
            return new byte[0];
        } else {
            XSSFWorkbook workBook = new XSSFWorkbook();
            XSSFSheet sheet = workBook.createSheet(numSheet.toString());
            Map<String, String> headerMap = (Map)datas.get(0);
            String[] fisrtHeaderKeyArray = (String[])headerMap.keySet().toArray(new String[headerMap.size()]);
            Iterator var7 = datas.iterator();

            while(var7.hasNext()) {
                Map<String, String> data = (Map)var7.next();
                int rowNum = datas.indexOf(data) + headerRowIndex;
                XSSFRow contentRow = sheet.createRow(rowNum);

                for(int cellIndex = 0; cellIndex < fisrtHeaderKeyArray.length; ++cellIndex) {
                    try {
                        contentRow.createCell(cellIndex).setCellValue((String)data.get(fisrtHeaderKeyArray[cellIndex]));
                    } catch (Exception var26) {
                        var26.printStackTrace();
                    }
                }
            }

            ByteArrayOutputStream byteArrayOutputStream = null;

            Object var31;
            try {
                byteArrayOutputStream = new ByteArrayOutputStream();
                workBook.write(byteArrayOutputStream);
                byte[] data = byteArrayOutputStream.toByteArray();
                byte[] var32 = data;
                return var32;
            } catch (Exception var27) {
                var27.printStackTrace();
                var31 = null;
            } finally {
                if (byteArrayOutputStream != null) {
                    try {
                        byteArrayOutputStream.close();
                    } catch (IOException var25) {
                        var25.printStackTrace();
                    }
                }

                try {
                    workBook.close();
                } catch (IOException var24) {
                    var24.printStackTrace();
                }

            }

            return (byte[])var31;
        }
    }

    public static byte[] exportMapDataExcel(Integer numSheet, Integer headerRowIndex, Map<String, Object> mapob, List<? extends Object> lobs, Object... objs) {
        Map<Integer, String> excelvalue_mapper = new HashMap();
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(numSheet.toString());
        XSSFRow headerRow = sheet.createRow(headerRowIndex);
        Integer headerCellIndex = 0;
        Iterator value = mapob.keySet().iterator();

        String key;
        while(value.hasNext()) {
            key = (String)value.next();
            headerRow.createCell(headerCellIndex).setCellValue(key);
            excelvalue_mapper.put(headerCellIndex, mapob.get(key).toString());
            if (headerCellIndex <= mapob.size()) {
                headerCellIndex = headerCellIndex + 1;
            }
        }

        value = null;
        key = null;
        Iterator var12 = lobs.iterator();

        XSSFRow contentRow;
        while(var12.hasNext()) {
            Object ob = var12.next();
            contentRow = sheet.createRow(lobs.indexOf(ob) + headerRowIndex + 1);

            for(int cellIndex = 0; cellIndex < headerCellIndex; ++cellIndex) {
                if (ob instanceof Map) {
                    try {
                        ob = (Map)ob;
                        String value = "";
                        if (excelvalue_mapper.containsKey(cellIndex)) {
                            key = (String)excelvalue_mapper.get(cellIndex);
                        }

                        if (((Map)ob).containsKey(key)) {
                            value = String.valueOf(((Map)ob).get(key));
                            if (value.equals("null") || value == null) {
                                value = "";
                            }
                        }

                        contentRow.createCell(cellIndex).setCellValue(String.valueOf(value));
                    } catch (Exception var28) {
                        var28.printStackTrace();
                    }
                }
            }
        }

        ByteArrayOutputStream baos = null;

        try {
            baos = new ByteArrayOutputStream();
            wb.write(baos);
            byte[] data = baos.toByteArray();
            byte[] var32 = data;
            return var32;
        } catch (Exception var26) {
            var26.printStackTrace();
            contentRow = null;
        } finally {
            if (baos != null) {
                try {
                    baos.close();
                } catch (IOException var25) {
                    var25.printStackTrace();
                }
            }

        }

        return (byte[])contentRow;
    }

    public static void limitExport(List<?> list) {
        if (org.apache.commons.collections.CollectionUtils.isNotEmpty(list)) {
            int size = list.size();
            if (size > 50000) {
                throw new BusinessException("一次性导出数据不能超过5万条");
            }
        }

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

    private void dealExcel(XSSFSheet sheet, Integer headerRowIndex, Map<String, Object> mapob, List<? extends Object> lobs, Object... objs) {
        Map<Integer, String> excelvalue_mapper = new HashMap();
        XSSFRow headerRow = sheet.createRow(headerRowIndex);
        Integer headerCellIndex = 0;
        Iterator var9 = mapob.keySet().iterator();

        while(var9.hasNext()) {
            String key = (String)var9.next();
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

    public static byte[] exportSimpleSheetExcel(Integer numSheet, Integer headerRowIndex, Map<String, Object> mapob, List<? extends Object> lobs, Object... objs) {
        Map<Integer, String> excelvalue_mapper = new HashMap();
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(numSheet.toString());
        XSSFRow headerRow = sheet.createRow(headerRowIndex);
        Integer headerCellIndex = 0;
        Iterator cellStyle = mapob.keySet().iterator();

        String dataFormat;
        while(cellStyle.hasNext()) {
            dataFormat = (String)cellStyle.next();
            headerRow.createCell(headerCellIndex).setCellValue(dataFormat);
            excelvalue_mapper.put(headerCellIndex, mapob.get(dataFormat).toString());
            if (headerCellIndex <= mapob.size()) {
                headerCellIndex = headerCellIndex + 1;
            }
        }

        cellStyle = null;
        dataFormat = null;
        XSSFCell cell = null;
        Iterator var13 = lobs.iterator();

        XSSFRow contentRow;
        while(var13.hasNext()) {
            Object ob = var13.next();
            contentRow = sheet.createRow(lobs.indexOf(ob) + headerRowIndex + 1);

            for(int cellIndex = 0; cellIndex <= headerCellIndex; ++cellIndex) {
                Field[] fields = ob.getClass().getDeclaredFields();
                Field[] var18 = fields;
                int var19 = fields.length;

                for(int var20 = 0; var20 < var19; ++var20) {
                    Field field = var18[var20];
                    field.setAccessible(true);
                    if (field.getName().equals(excelvalue_mapper.get(cellIndex))) {
                        try {
                            cell = contentRow.createCell(cellIndex);
                            if (field.getGenericType().toString().equals("class java.math.BigDecimal")) {
                                XSSFCellStyle cellStyle = wb.createCellStyle();
                                XSSFDataFormat dataFormat = wb.createDataFormat();
                                cellStyle.setDataFormat(dataFormat.getFormat("0.00"));
                                cell.setCellValue(((BigDecimal)field.get(ob)).doubleValue());
                            } else {
                                cell.setCellValue(String.valueOf(field.get(ob)));
                            }
                        } catch (Exception var32) {
                            var32.printStackTrace();
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
            byte[] var39 = data;
            return var39;
        } catch (Exception var33) {
            var33.printStackTrace();
            contentRow = null;
        } finally {
            if (baos != null) {
                try {
                    baos.close();
                } catch (IOException var31) {
                    var31.printStackTrace();
                }
            }

        }

        return (byte[])contentRow;
    }*/
}
