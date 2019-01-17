import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Map.Entry;

public class ExcelUtil2 {

    /**
     *
     * 将对象数组转换成excel<br/>
     * <dependency>
     * 	<groupId>org.apache.poi</groupId>
     * 	<artifactId>poi-ooxml</artifactId>
     * 	<version>3.17</version>
     * </dependency>
     * <dependency>
     * 	<groupId>commons-beanutils</groupId>
     * 	<artifactId>commons-beanutils</artifactId>
     * 	<version>1.9.3</version>
     * </dependency>
     *
     * @param pojoList  对象数组
     * @param out       输出流
     * @param alias     指定对象属性别名，生成列名和列顺序Map<"类属性名","列名">
     * @param utilExcel 表头对象
     * @throws Exception
     */
    public static <T> void pojo2Excel(List<T> pojoList, OutputStream out, LinkedHashMap<String, String> alias, UtilExcel utilExcel) throws Exception {
        if (utilExcel == null) throw new Exception("UtilExcel 对象为空");

        //创建一个工作簿
        XSSFWorkbook wb = new XSSFWorkbook();
        //创建一个表
        XSSFSheet sheet = wb.createSheet();
        //创建第一行，作为表名
        XSSFRow row = sheet.createRow(0);// 这个方法感觉是直接跳到对应行的
        XSSFCell cell = row.createCell(0);
        cell.setCellValue(utilExcel.getTableHeadName());
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, alias.size() - 1));

        // 在第一行插入列名
        insertColumnName(utilExcel.getDataStarRow() - 1, sheet, alias);

        // 从第指定行开始插入数据
        insertColumnDate(utilExcel.getDataStarRow(), pojoList, sheet, alias);

        // 输出表格文件
        try {
            wb.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            wb.close();
        }
    }

    /**
     * 将对象数组转换成excel，并增加复杂表头，除去表名和显示列的名称的那一行
     *
     * @param pojoList      对象数组
     * @param out           输出流
     * @param alias         指定对象属性别名，生成列名和列顺序Map<"类属性名","列名">
     * @param utilExcel     表头对象
     * @param mergeDataList 合并行中的所有数据，包括不和并的
     * @throws Exception
     */
    public static <T> void pojo2Excel(List<T> pojoList, OutputStream out, LinkedHashMap<String, String> alias, UtilExcel utilExcel, List<MergeData> mergeDataList) throws Exception {
        if (utilExcel == null) throw new Exception("UtilExcel 对象为空");
        //创建一个工作簿
        XSSFWorkbook wb = new XSSFWorkbook();
        //创建一个表
        XSSFSheet sheet = wb.createSheet();
        //创建第一行，作为表名
        XSSFRow row = sheet.createRow(0);// 后面不需要再次调用该方法，应该是使用该方法 可以独立设置该行的样式
        XSSFCell cell = row.createCell(0);
//        设置居中
//        设置表头文字格式
//        XSSFCellStyle cellStyle = wb.createCellStyle();
//        XSSFFont font = wb.createFont();
//        font.setFontName("宋体");
//        font.setFontHeightInPoints((short) 36);
//        cellStyle.setFont(font);
//        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cell.setCellValue(utilExcel.getTableHeadName());
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, alias.size() + 1));
        // 插入复杂表头
        XSSFRow rowTable = sheet.createRow(utilExcel.getDataStarRow() - 3);
        for (MergeData mergeData : mergeDataList) {
            sheet.addMergedRegion(new CellRangeAddress(mergeData.getStartRow(), mergeData.getEndRow(), mergeData.getStartCol(), mergeData.getEndCol()));
//            插入数据
            XSSFCell tableCellValue = rowTable.createCell(mergeData.getStartCol());
            tableCellValue.setCellValue(mergeData.getName());
            // 这里可以对单元格做样式处理
        }

        // 在第一行插入列名
        insertColumnName(utilExcel.getDataStarRow() - 1, sheet, alias);

        // 从第指定行开始插入数据
        insertColumnDate(utilExcel.getDataStarRow(), pojoList, sheet, alias);

        // 输出表格文件
        try {
            wb.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            wb.close();
        }
    }


    /**
     * 将excel表转换成指定类型的对象数组
     *
     * @param claz  类型
     * @param alias 列别名,格式要求：Map<"列名","类属性名">
     * @param param 指定第几行行为字段名，第一行为1
     * @return
     * @throws IOException
     * @throws IllegalArgumentException
     * @throws IllegalAccessException
     * @throws SecurityException
     * @throws NoSuchFieldException
     * @throws InstantiationException
     * @throws InvocationTargetException
     */
    public static <T> List<T> excel2Pojo(InputStream inputStream, Class<T> claz, LinkedHashMap<String, String> alias, Integer param) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook(inputStream);
        try {
            XSSFSheet sheet = wb.getSheetAt(0);

            //生成属性和列对应关系的map，Map<类属性名，对应一行的第几列>
            Map<String, Integer> propertyMap = generateColumnPropertyMap(sheet, alias, param);
            //根据指定的映射关系进行转换
            List<T> pojoList = generateList(sheet, propertyMap, claz, param);
            return pojoList;
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        } finally {
            wb.close();
        }
    }


    /**
     * 将对象数组转换成excel
     *
     * @param pojoList 对象数组
     * @param out      输出流
     * @param alias    指定对象属性别名，生成列名和列顺序
     * @throws Exception
     */
    public static <T> void pojo2Excel(List<T> pojoList, OutputStream out, LinkedHashMap<String, String> alias) throws Exception {
        //获取类名作为标题
        String headLine = "";
        if (pojoList.size() > 0) {
            Object pojo = pojoList.get(0);
            Class<? extends Object> claz = pojo.getClass();
            headLine = claz.getName();
            pojo2Excel(pojoList, out, alias, new UtilExcel(headLine, 1));
        }
    }

    /**
     * 将对象数组转换成excel,列名为对象属性名
     *
     * @param pojoList 对象数组
     * @param out      输出流
     * @param headLine 表标题
     * @throws Exception
     */
    public static <T> void pojo2Excel(List<T> pojoList, OutputStream out, String headLine) throws Exception {
        //获取类的属性作为列名
        LinkedHashMap<String, String> alias = new LinkedHashMap<String, String>();
        if (pojoList.size() > 0) {
            Object pojo = pojoList.get(0);
            Field[] fields = pojo.getClass().getDeclaredFields();
            String[] name = new String[fields.length];
            Field.setAccessible(fields, true);
            for (int i = 0; i < name.length; i++) {
                name[i] = fields[i].getName();
                alias.put(isNull(name[i]).toString(), isNull(name[i]).toString());
            }
            pojo2Excel(pojoList, out, alias, new UtilExcel(headLine, 1));
        }
    }

    /**
     * 将对象数组转换成excel，列名默认为对象属性名，标题为类名
     *
     * @param pojoList 对象数组
     * @param out      输出流
     * @throws Exception
     */
    public static <T> void pojo2Excel(List<T> pojoList, OutputStream out) throws Exception {
        //获取类的属性作为列名
        LinkedHashMap<String, String> alias = new LinkedHashMap<String, String>();
        //获取类名作为标题
        String headLine = "";
        if (pojoList.size() > 0) {
            Object pojo = pojoList.get(0);
            Class<? extends Object> claz = pojo.getClass();
            headLine = claz.getName();
            Field[] fields = claz.getDeclaredFields();
            String[] name = new String[fields.length];
            Field.setAccessible(fields, true);
            for (int i = 0; i < name.length; i++) {
                name[i] = fields[i].getName();
                alias.put(isNull(name[i]).toString(), isNull(name[i]).toString());
            }
            pojo2Excel(pojoList, out, alias, new UtilExcel(headLine, 1));
        }
    }

    /**
     * 此方法作用是创建表头的列名
     *
     * @param alias  要创建的表的列名与实体类的属性名的映射集合
     * @param rowNum 指定行创建列名
     * @return
     */
    private static void insertColumnName(int rowNum, XSSFSheet sheet, Map<String, String> alias) {
        XSSFRow row = sheet.createRow(rowNum);
        //列的数量
        int columnCount = 0;

        Set<Entry<String, String>> entrySet = alias.entrySet();

        for (Entry<String, String> entry : entrySet) {
            //创建第一行的第columnCount个格子
            XSSFCell cell = row.createCell(columnCount++);
            //将此格子的值设置为alias中的键名
            cell.setCellValue(isNull(entry.getValue()).toString());
        }
    }

    /**
     * 从指定行开始插入数据
     *
     * @param beginRowNum 开始行
     * @param models      对象数组
     * @param sheet       表
     * @param alias       列别名
     * @throws Exception
     */
    private static <T> void insertColumnDate(int beginRowNum, List<T> models, XSSFSheet sheet, Map<String, String> alias) throws Exception {
        for (T model : models) {
            //创建新的一行
            XSSFRow rowTemp = sheet.createRow(beginRowNum++);
            //获取列的迭代
            Set<Entry<String, String>> entrySet = alias.entrySet();

            //从第0个格子开始创建
            int columnNum = 0;
            for (Entry<String, String> entry : entrySet) {
                //获取属性值
                String property = BeanUtils.getProperty(model, entry.getKey());
                //创建一个格子
                XSSFCell cell = rowTemp.createCell(columnNum++);
                // 得知string可以转化的类型
                if (isDouble(property)){
                    cell.setCellValue(Double.valueOf(property));
                }else if (isInt(property)){
                    cell.setCellValue(Integer.valueOf(property));
                }else if (isDateAndTime(property)){
                    // 只对日期加time的做转化
                    SimpleDateFormat formatter;
                    if (property.indexOf("-") >= 1){
                        formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                    }else if(property.indexOf("/") >= 1){
                        formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                    }else if(property.indexOf(".") >= 1){
                        formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                    }else {
                        formatter = new SimpleDateFormat("yyyyMMdd HH:mm:ss");
                    }
                    Date date = formatter.parse(property);
                    cell.setCellValue(date);
                }else {
                    cell.setCellValue(property);
                }
            }
        }
    }

    //判断是否为空，若为空设为""
    private static Object isNull(Object object) {
        if (object != null) {
            return object;
        } else {
            return "";
        }
    }

    /**
     * 生成一个属性-列的对应关系的map
     *
     * @param sheet 表
     * @param alias 别名
     * @return
     */
    private static Map<String, Integer> generateColumnPropertyMap(XSSFSheet sheet, LinkedHashMap<String, String> alias, Integer param) {
        Map<String, Integer> propertyMap = new HashMap<>();

        if (param == null || param < 0) param = 1;
        XSSFRow propertyRow = sheet.getRow(param);
        short firstCellNum = propertyRow.getFirstCellNum();
        short lastCellNum = propertyRow.getLastCellNum();

        for (int i = firstCellNum; i < lastCellNum; i++) {
            Cell cell = propertyRow.getCell(i);
            if (cell == null) {
                continue;
            }
            //列名
            String cellValue = cell.getStringCellValue();
            //对应属性名
            String propertyName = alias.get(cellValue);
            propertyMap.put(propertyName, i);
        }
        return propertyMap;
    }

    /**
     * 根据指定关系将表数据转换成对象数组
     *
     * @param sheet       表
     * @param propertyMap 属性映射关系Map<"属性名",一行第几列>
     * @param claz        类类型
     * @return
     * @throws InstantiationException
     * @throws IllegalAccessException
     * @throws InvocationTargetException
     */
    private static <T> List<T> generateList(XSSFSheet sheet, Map<String, Integer> propertyMap, Class<T> claz, Integer param) throws Exception {
        if (param == null || param < 0) param = 1;
        //对象数组
        List<T> pojoList = new ArrayList<>();
        int index = 0;
        for (Row row : sheet) {
            //跳过标题和列名
            if (row.getRowNum() < param + 1) {
                continue;
            }
            T instance = claz.newInstance();
            Set<Entry<String, Integer>> entrySet = propertyMap.entrySet();
            for (Entry<String, Integer> entry : entrySet) {
                /*
                 * CellTypeEnum        类型        值
                 * NUMERIC             数值型      0
                 * STRING              字符串型     1
                 * FORMULA             公式型      2
                 * BLANK               空值        3
                 * BOOLEAN             布尔型      4
                 * ERROR               错误        5
                 *
                 * 4.0以上将会移除 替换为getCellType
                 * */
                // 获取此行指定列的值,即为属性对应的值
                switch (row.getCell(entry.getValue()).getCellTypeEnum()) {
                    case _NONE:
                        System.out.println("****************************不知道的类型*********************************");
                        throw new Exception("第" + index + "行【" + row.getCell(entry.getValue()) + "】导入数据异常");
                    case BLANK:
                        BeanUtils.setProperty(instance, entry.getKey(), null);
                        break;
                    case NUMERIC:
                        int pInt = (int) row.getCell(entry.getValue()).getNumericCellValue();
                        BeanUtils.setProperty(instance, entry.getKey(), pInt);
                        break;
                    case STRING:
                        String pString = row.getCell(entry.getValue()).getStringCellValue();
                        BeanUtils.setProperty(instance, entry.getKey(), pString);
                        break;
                    case FORMULA:
                        System.out.println("****************************【公式】没做处理");
                        break;
                    case BOOLEAN:
                        boolean pBoolean = row.getCell(entry.getValue()).getBooleanCellValue();
                        BeanUtils.setProperty(instance, entry.getKey(), pBoolean);
                        break;
                    case ERROR:
                        System.out.println("****************************error*********************************");
                        throw new Exception("第" + index + "行【" + row.getCell(entry.getValue()) + "】导入数据异常");
                }
            }
            pojoList.add(instance);
            index++;
        }
        return pojoList;
    }

    /**
     * 将excel表转换成指定类型的对象数组，列名即作为对象属性
     *
     * @param claz 类型
     * @return
     * @throws IOException
     * @throws InstantiationException
     * @throws SecurityException
     * @throws NoSuchFieldException
     * @throws IllegalAccessException
     * @throws IllegalArgumentException
     * @throws InvocationTargetException
     */
    public static <T> List<T> excel2Pojo(InputStream inputStream, Class<T> claz) throws IllegalArgumentException, IllegalAccessException, NoSuchFieldException, SecurityException, InstantiationException, IOException, InvocationTargetException {
        LinkedHashMap<String, String> alias = new LinkedHashMap<String, String>();
        Field[] fields = claz.getDeclaredFields();
        for (Field field : fields) {
            alias.put(field.getName(), field.getName());
        }
        List<T> pojoList = excel2Pojo(inputStream, claz, alias, 1);
        return pojoList;
    }


    /**
     * String可以转化的类型判断
     *
     * @param str
     */
//    是否为浮点数
    private static boolean isDouble(String str) {
        return str.matches("^[-+]?[1-9][0-9]*\\.?[0-9]+$");
    }
//    是否为整数
    private static boolean isInt(String str) {
        return str.matches("^[-+]?[1-9]\\d*$");
    }

//    必须日期加时间 [2018-02-14 00:00:00] 使用反向引用进行简化，年份0001-9999，格式yyyy-MM-dd或yyyy-M-d，连字符可以没有或是“-”、“/”、“.”之一。
    private static boolean isDateAndTime(String str) {
        return str.matches("^(?:(?!0000)[0-9]{4}([-/.]?)(?:(?:0?[1-9]|1[0-2])\\1(?:0?[1-9]|1[0-9]|2[0-8])|(?:0?[13-9]|1[0-2])\\1(?:29|30)|(?:0?[13578]|1[02])\\1(?:31))|(?:[0-9]{2}(?:0[48]|[2468][048]|[13579][26])|(?:0[48]|[2468][048]|[13579][26])00)([-/.]?)0?2\\2(?:29))\\s+([01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]$");
    }
//    必须日期加时间 [2018-02-14] 使用反向引用进行简化，年份0001-9999，格式yyyy-MM-dd或yyyy-M-d，连字符可以没有或是“-”、“/”、“.”之一。
    private static boolean isDate(String str) {
        return str.matches("^(?:(?!0000)[0-9]{4}([-/.]?)(?:(?:0?[1-9]|1[0-2])([-/.]?)(?:0?[1-9]|1[0-9]|2[0-8])|(?:0?[13-9]|1[0-2])([-/.]?)(?:29|30)|(?:0?[13578]|1[02])([-/.]?)31)|(?:[0-9]{2}(?:0[48]|[2468][048]|[13579][26])|(?:0[48]|[2468][048]|[13579][26])00)([-/.]?)0?2([-/.]?)29)$");
    }
}

class UtilExcel {
    private String tableHeadName;// 表头名称
    private int dataStarRow;// 插入数据开始的row

    UtilExcel() {
        this.tableHeadName = "export excel";
        this.dataStarRow = 1;
    }

    public UtilExcel(String tableHeadName, int dataStarRow) {
        this.tableHeadName = tableHeadName;
        this.dataStarRow = dataStarRow;
    }

    public String getTableHeadName() {
        return tableHeadName;
    }

    public void setTableHeadName(String tableHeadName) {
        this.tableHeadName = tableHeadName;
    }

    public int getDataStarRow() {
        return dataStarRow;
    }

    public void setDataStarRow(int dataStarRow) {
        this.dataStarRow = dataStarRow;
    }
}

class MergeData {
    private String name;
    private int startRow;
    private int endRow;
    private int startCol;
    private int endCol;

    public MergeData() {
    }

    public MergeData(String name, int startRow, int endRow, int startCol, int endCol) {
        this.name = name;
        this.startRow = startRow;
        this.endRow = endRow;
        this.startCol = startCol;
        this.endCol = endCol;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getStartRow() {
        return startRow;
    }

    public void setStartRow(int startRow) {
        this.startRow = startRow;
    }

    public int getEndRow() {
        return endRow;
    }

    public void setEndRow(int endRow) {
        this.endRow = endRow;
    }

    public int getStartCol() {
        return startCol;
    }

    public void setStartCol(int startCol) {
        this.startCol = startCol;
    }

    public int getEndCol() {
        return endCol;
    }

    public void setEndCol(int endCol) {
        this.endCol = endCol;
    }
}