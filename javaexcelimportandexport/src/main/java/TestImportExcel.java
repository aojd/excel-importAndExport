import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
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
    public static void test() throws Exception {
        //指定输入文件
        FileInputStream fis = new FileInputStream("D:\\data\\test.xls");
        //指定每列对应的类属性
        LinkedHashMap<String, String> alias = new LinkedHashMap<>();
        alias.put("姓名", "name");
        alias.put("年龄", "age");
        //转换成指定类型的对象数组
        List<TestUser> pojoList = ExcelUtil2.excel2Pojo(fis, TestUser.class, alias, 0);

        System.out.println(pojoList.get(0).getName());
//        logger.info(pojoList.toString());


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
//        list.add(new MergeData("end", 1, 2, 2, 3));
        ExcelUtil2.pojo2Excel(pojoList, fos, aliasE, new UtilExcel("table",4),list);
//        ExcelUtil2.pojo2Excel(pojoList, fos, aliasE, new UtilExcel("table",2));
    }

    public static void main(String[] args) {
        try {
            test();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
