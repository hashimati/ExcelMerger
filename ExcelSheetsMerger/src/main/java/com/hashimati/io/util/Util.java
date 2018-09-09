package com.hashimati.io.util;

public class Util
{

    public static String whatIsYourOS(){

        return System.getProperty("os.name").toLowerCase();

    }

    public static boolean isOsWin() {
        return whatIsYourOS().contains("windows");

    }
}
