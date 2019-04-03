package com.yang.serialport.utils;

import java.io.*;

/**
 * Created by liu_changshi on 2019/4/2.
 */
public class copyExcel {
    public static void copy2(String readfile,String writeFile) {
        try {
            FileInputStream input = new FileInputStream(readfile);
            FileOutputStream output = new FileOutputStream(writeFile);
            int read = input.read();
            while ( read != -1 ) {
                output.write(read);
                read = input.read();
            }
            input.close();
            output.close();
        } catch (IOException e) {
            System.out.println(e);
        }
    }
    /*public static void main(String[] args) {
        String path1 = "C:/Users/liu_changshi/Desktop/真空阀门成品检验报告(中车50-2专用)-仪器专用.xls";//源文件
        String dir = "C:/Users/liu_changshi/Desktop/目标文件夹/";//指定的文件夹
        String fileName = "2.xls";//指定的重命名
        copy2(path1, dir+fileName);
    }*/


}
