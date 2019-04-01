package com.yang.serialport.utils;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class writeExcel{
    public static int[] getRowNumAndColumnNum(String location){



    }

    public static void setExcel(String pathName,int rowNum,int columnNum,String content) throws IOException {
        FileInputStream fs=new FileInputStream(pathName);  //获取d://test.xls
        POIFSFileSystem ps=new POIFSFileSystem(fs);  //使用POI提供的方法得到excel的信息
        HSSFWorkbook wb=new HSSFWorkbook(ps);
        HSSFSheet sheet=wb.getSheetAt(0);  //获取到工作表，因为一个excel可能有多个工作表

        FileOutputStream out=new FileOutputStream(pathName);  //向d://test.xls中写数据
        HSSFRow row=sheet.getRow(rowNum); //在现有行号后追加数据
        row.createCell(columnNum).setCellValue(content); //设置第一个（从0开始）单元格的数据



        out.flush();
        wb.write(out);
        out.close();

    }

    /*public static void main(String[] args) throws Exception {
        String path = "C://Users/liu_changshi/desktop/真空阀门成品检验报告(中车50-2专用)-仪器专用.xls";
        int a = 10;
        int b = 5;
        String z = "gggg";
        writeExcel.setExcel(path,a,b,z);

    }*/
}   