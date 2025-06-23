package com.example;

import java.util.ArrayList;
import java.util.HashMap;
// import com.example.ExcelUtil;
import java.util.List;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

/**
 * Hello world!
 *
 */
public class App {
    static ArrayList<String> list = new ArrayList<String>();

    static HashMap<String, Integer> map = new HashMap<String, Integer>();

    public static void main(String[] args) {
        System.out.println("Hello World!\r\n");
        Logger logger = LogManager.getLogger(App.class);
        logger.info("Hello ---World!");
        String c = logger.toString();
        System.out.println("c的结果是:" + c);
        // ExcelUtil.list
        // ExcelUtil util = new ExcelUtil();
        // // List<List<List<String>>> allSheet = util.readAllSheet("D:\\test.xlsx");
        // List<List<String>> firstSheet = util.readFirstSheet("D:\\test.xlsx");

        String excelPath = "F:\\Java_Projuct\\arryle\\example1.xlsx";
        List<List<String>> excelData = ExcelUtil.readFirstSheet(excelPath);
        System.out.println("excel中第1个sheet的内容:" + excelData);
        for (List<String> row : excelData) {
            for (String cell : row) {
                System.out.println(cell);
            }
        }

        // excelData.

        list.add("Hello");
        list.add("World");

        boolean a = list.contains("qq");
        System.out.println("a的结果是:" + a);

        map.put("Hello", 1);
        map.put("World", 2);

        // map.replace

        System.out.println("Hello World!");
        System.out.println(list.get(0));
        System.out.println(list.get(1));

        // 遍历ArrayList
        for (String str : list) {
            System.out.println(str);
        }

        for (String key : map.keySet()) {
            System.out.println(key + " : " + map.get(key));
        }
    }
}
