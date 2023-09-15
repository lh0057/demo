package com.poniard;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class WorkBookUtils {
//"C:\\Users\\lh\\Desktop\\闲言观众统计表03.xls"
    public static void getWorkBook(String path){
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(new File(path));
            Sheet sheet = workbook.getSheetAt(0);
            int len = sheet.getLastRowNum();
            for(int i=0;i<=len;i++){
                Row row = sheet.getRow(i);
                if (row == null){
                    continue;
                }
                int size = row.getLastCellNum();
                for (int j=0;j<size;j++){
                    Cell cell = row.getCell(j);
                    if (cell == null){
                        continue;
                    }
                    int type = cell.getCellType();
                    if(type == 0){
                        System.out.print(cell.getNumericCellValue());
                    }else if (type == 1){
                        System.out.print(cell.getStringCellValue().trim().replaceAll("\\s",""));
                    }

                    if (j<size-1){
                        System.out.print(" | ");
                    }
                }
                System.out.println("");
            }

            System.out.println("文件读取完毕");
        } catch (IOException e) {
            throw new RuntimeException(e);
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }
    }

    public static void write03(){
        Workbook workbook = new HSSFWorkbook();
        //2.创建 一个工作表
        Sheet sheet = workbook.createSheet("闲言粉丝统计表");
        //3.创建一行
        Row row1 = sheet.createRow(0);
        //4.创建一个单元格
        //(1,1)
        Cell cell1 = row1.createCell(0);
        cell1.setCellValue("今日新增观众");
        //(1,2)
        Cell cell2 = row1.createCell(1);
        cell2.setCellValue(666);

        //创建第二行
        Row row2 = sheet.createRow(1);
        //(2,1)
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");
        //(2,2)
        Cell cell22 = row2.createCell(1);
        Date date = new Date();
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy年MM月dd日 hh时mm分ss秒");
        String datetime = simpleDateFormat.format(date);

        cell22.setCellValue(datetime);

        //生成一张表(IO流)，03版本就是使用xls结尾
        writeFile(workbook,"C:\\Users\\lh\\Desktop\\闲言观众统计表03.xls");
        System.out.println("文件生成完毕");
    }

    public static void write07() throws Exception {
        //1.创建一个工作簿
        Workbook workbook = new XSSFWorkbook();
        //2.创建 一个工作表
        Sheet sheet = workbook.createSheet("闲言粉丝统计表");
        //3.创建一行
        Row row1 = sheet.createRow(0);
        //4.创建一个单元格
        //(1,1)
        Cell cell1 = row1.createCell(0);
        cell1.setCellValue("今日新增观众");
        //(1,2)
        Cell cell2 = row1.createCell(1);
        cell2.setCellValue(666);

        //创建第二行
        Row row2 = sheet.createRow(1);
        //(2,1)
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");
        //(2,2)
        Cell cell22 = row2.createCell(1);
        Date date = new Date();
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy年MM月dd日 hh时mm分ss秒");
        String datetime = simpleDateFormat.format(date);
        cell22.setCellValue(datetime);

        //生成一张表(IO流)，03版本就是使用xlsx结尾
        writeFile(workbook,"C:\\Users\\lh\\Desktop\\闲言观众统计表03.xlsx");
        System.out.println("文件生成完毕");
    }

    public static void writeFile(Workbook workbook, String path){
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(path);
            //输出
            workbook.write(fos);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            try {
                //关闭流
                fos.close();
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
    }

    public static void writeFile(Workbook workbook, File file){
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(file);
            //输出
            workbook.write(fos);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            try {
                //关闭流
                fos.close();
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
    }


}
