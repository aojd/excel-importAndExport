package top.aojd.bookstore.util.toolkit;

import org.apache.commons.beanutils.PropertyUtils;

import java.io.*;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;

/**
 * The Class :
 * Author: AOJD
 * Data: 2019/1/9
 * Time: 16:12
 * Version:
 */
public class TestImportExcel {
    public static void testSheetList() throws Exception {
        //指定输入文件
        FileInputStream fis = new FileInputStream("D:\\data\\test2.xls");

        List<SheetImport> listImport = new ArrayList<>();
        //指定每列对应的类属性
        LinkedHashMap<String, String> alias = new LinkedHashMap<>();
        alias.put("姓名", "name");
        alias.put("年龄", "age");
        SheetImport sheetImport = new SheetImport();
        sheetImport.setAlias(alias);
        sheetImport.setClaz(TestUser.class);
        sheetImport.setParam(0);
        listImport.add(sheetImport);
        LinkedHashMap<String, String> alias2 = new LinkedHashMap<>();
        alias2.put("姓名", "name");
        alias2.put("年龄", "age");
        alias2.put("日期", "dat");
        SheetImport sheetImport2 = new SheetImport();
        sheetImport2.setAlias(alias2);
        sheetImport2.setClaz(TestUserSS.class);
        sheetImport2.setParam(0);
        listImport.add(sheetImport2);
        //转换成指定类型的对象数组
        List<?> pojoList = ExcelUtil2.excel2PojoSheetList(listImport, fis);
        List<TestUser> testUserList = new ArrayList<>();
        for (TestUser te : (List<TestUser>) pojoList.get(0)) {
            TestUser testUser = new TestUser();
            PropertyUtils.copyProperties(testUser, te);
            testUserList.add(testUser);
        }

        List<TestUserSS> testUserSSList = new ArrayList<>();
        for (TestUserSS te : (List<TestUserSS>) pojoList.get(1)) {
            TestUserSS testUseSS = new TestUserSS();
            PropertyUtils.copyProperties(testUseSS, te);
            testUserSSList.add(testUseSS);
        }
//        logger.info(pojoList.toString());


//        *****************************************************Export excel******************************************************************

        //将生成的excel转换成文件，还可以用作文件下载
        File file = new File("D:\\data\\testExport2.xls");
        FileOutputStream fos = new FileOutputStream(file);

        List<SheetExport> sheetExport = new ArrayList<>();

        SheetExport us = new SheetExport();
        LinkedHashMap<String, String> aliaEp = new LinkedHashMap<>();
        aliaEp.put("name", "姓名");
        aliaEp.put("age", "年龄");
        us.setAlias(aliaEp);
        us.setSheetName("第一");
        us.setPojoList(testUserList);
        us.setUtilExcel(new UtilExcel("table", 1));
        sheetExport.add(us);

        SheetExport usS = new SheetExport();
        LinkedHashMap<String, String> aliasEx = new LinkedHashMap<>();
        aliasEx.put("name", "姓名");
        aliasEx.put("age", "年龄");
        aliasEx.put("dat", "日期");
        usS.setAlias(aliasEx);
        usS.setSheetName("sheet name");
        usS.setPojoList(testUserSSList);
        List<MergeData> lisor = new ArrayList<>();
        lisor.add(new MergeData("start", 1, 2, 0, 2));
        usS.setMergeDataList(lisor);
        usS.setUtilExcel(new UtilExcel("table", 1));
        sheetExport.add(usS);

        ExcelUtil2.pojo2ExcelSheetList(sheetExport, fos);
    }


    public static void test() throws Exception {
        // 指定输入文件
        FileInputStream fis = new FileInputStream("D:\\data\\test2.xls");
        // 指定每列对应的类属性
        LinkedHashMap<String, String> alias = new LinkedHashMap<>();
        alias.put("姓名", "name");
        alias.put("年龄", "age");
        // 转换成指定类型的对象数组
        List<TestUser> pojoList = ExcelUtil2.excel2Pojo(fis, TestUser.class, alias, 0);

        System.out.println(pojoList.get(0).getName());
       // logger.info(pojoList.toString());


//        *****************************************************Export excel******************************************************************

        //将生成的excel转换成文件，还可以用作文件下载
        File file = new File("D:\\data\\testExport.xls");
        FileOutputStream fos = new FileOutputStream(file);

        //对象集合
        List<TestUser> pojoExport = new ArrayList<>();
        for (int i = 0; i < 5; i++) {
            TestUser user = new TestUser();
            user.setName("老李");
            user.setAge(50);
            pojoExport.add(user);
        }
        //设置属性别名（列名）
        LinkedHashMap<String, String> aliasE = new LinkedHashMap<>();
        aliasE.put("name", "姓名");
        aliasE.put("age", "年龄");
        //标题
        String headLine = "用户表";
        List<MergeData> list = new ArrayList<>();
        list.add(new MergeData("start", 1, 2, 0, 2));
        // list.add(new MergeData("end", 1, 2, 2, 3));
        ExcelUtil2.pojo2Excel(pojoList, fos, aliasE, new UtilExcel("table", 4), list);
        // ExcelUtil2.pojo2Excel(pojoList, fos, aliasE, new UtilExcel("table",2));
    }

    public static void main(String[] args) {
        try {
            testSheetList();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
