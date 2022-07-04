package com.yyyang;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.beans.Transient;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

public class ExcelWriteTest {

    String PATH = "E:\\EclipseProject\\mavenProject";

    @Test
    public void testWrite03() throws Exception {
        //1.创建一个工作簿
        Workbook workbook = new HSSFWorkbook();
        //2.创建一个工作表
        Sheet sheet = workbook.createSheet("续费统计表");
        //3.创建一个行
        Row row1 = sheet.createRow(0);
        //4.创建一个单元格（1,1） (1,2)
        Cell cell_11 = row1.createCell(0);
        cell_11.setCellValue("今日新增续费");
        Cell cell_12 = row1.createCell(1);
        cell_12.setCellValue("扣费时间");

        //第二行
        Row row2 = sheet.createRow(1);

        Cell cell_21 = row2.createCell(0);
        cell_21.setCellValue("yyyang");
        Cell cell_22 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-mm-dd HH:mm:ss");
        cell_22.setCellValue(time);

        //生成一张表（IO 流） 03版本使用xls结尾
        FileOutputStream outputStream = new FileOutputStream(PATH + "续费统计表03.xls");
        //输出
        workbook.write(outputStream);
        //关闭流
        outputStream.close();

        System.out.println("续费统计表03输出完毕！");

    }

    @Test
    public void testWrite07() throws Exception {
        //1.创建一个工作簿
        Workbook workbook = new XSSFWorkbook();
        //2.创建一个工作表
        Sheet sheet = workbook.createSheet("续费统计表");
        //3.创建一个行
        Row row1 = sheet.createRow(0);
        //4.创建一个单元格（1,1） (1,2)
        Cell cell_11 = row1.createCell(0);
        cell_11.setCellValue("今日新增续费");
        Cell cell_12 = row1.createCell(1);
        cell_12.setCellValue("扣费时间");

        //第二行
        Row row2 = sheet.createRow(1);

        Cell cell_21 = row2.createCell(0);
        cell_21.setCellValue("yyyang");
        Cell cell_22 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-mm-dd HH:mm:ss");
        cell_22.setCellValue(time);

        //生成一张表（IO 流） 07版本使用xlsx结尾
        FileOutputStream outputStream = new FileOutputStream(PATH + "续费统计表07.xlsx");
        //输出
        workbook.write(outputStream);
        //关闭流
        outputStream.close();

        System.out.println("续费统计表07输出完毕！");

    }

    @Test
    public void testWrite03BigData() throws Exception {
        //时间
        long begin = System.currentTimeMillis();

        //创建一个工作簿
        Workbook workbook = new HSSFWorkbook();

        //创建一张表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNum=0; rowNum<65536; rowNum++){
            Row row = sheet.createRow(rowNum);
            for (int cellNum=0; cellNum<10; cellNum++){
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("over");
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "big03.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();

        long end = System.currentTimeMillis();
        System.out.println((end-begin)/1000);

    }


    @Test
    public void testWrite07BigData() throws Exception {
        //时间
        long begin = System.currentTimeMillis();

        //创建一个工作簿
        Workbook workbook = new XSSFWorkbook();

        //创建一张表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNum=0; rowNum<100000; rowNum++){
            Row row = sheet.createRow(rowNum);
            for (int cellNum=0; cellNum<10; cellNum++){
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("over");
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "big07.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();

        long end = System.currentTimeMillis();
        System.out.println((end-begin)/1000);

    }

    @Test
    public void testWrite07BigDataS() throws Exception {
        //时间
        long begin = System.currentTimeMillis();

        //创建一个工作簿
        Workbook workbook = new SXSSFWorkbook();

        //创建一张表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNum=0; rowNum<100000; rowNum++){
            Row row = sheet.createRow(rowNum);
            for (int cellNum=0; cellNum<10; cellNum++){
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("over");
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "big07s.xlsx");
        workbook.write(fileOutputStream);
        //清楚零时文件
        ((SXSSFWorkbook) workbook).dispose();
        fileOutputStream.close();

        long end = System.currentTimeMillis();
        System.out.println((end-begin)/1000);

    }

}
