package com.poniard;

import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;

public class Main {
    public static void main(String[] args) {
//        WorkBookUtils.write03();
//        try {
//            WorkBookUtils.write07();
//        } catch (Exception e) {
//            throw new RuntimeException(e);
//        }
        WorkBookUtils.getWorkBook("C:\\Users\\lh\\Desktop\\生产与svn比对.xlsx");
    }

}