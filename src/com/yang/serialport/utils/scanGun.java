package com.yang.serialport.utils;

import java.util.Scanner;

/**
 * Created by liu_changshi on 2019/4/1.
 */
public class scanGun {
    public static String getscanGunData(){
        Scanner sr = new Scanner(System.in);
        String name = sr.nextLine();
        return name;
    }


}
