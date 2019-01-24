package com.yneusoft.euexcel.test.demo;

import cn.afterturn.easypoi.excel.annotation.Excel;
import com.yneusoft.euexcel.test.entity.template.StudentTemplate;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Map导入
 * @author 易用软件-朱新旺(zhuxinwang@aliyun.com)
 * @date 2019/1/23 0023 10:02
 */
public class Demo04 {


    static String PATH = "D:\\test\\workbook56.xls";

    public static void main(String[] args) throws IOException, IllegalAccessException, InstantiationException, ParseException {

        long startTime = System.currentTimeMillis();

        List<StudentTemplate> studentTemplateList = null;


        InputStream is = new FileInputStream(PATH);
        Workbook workbook = new HSSFWorkbook(is);
        studentTemplateList = populate(workbook,StudentTemplate.class);

        System.out.println("运行总时间：" + (System.currentTimeMillis() - startTime));

        System.out.println(Arrays.toString(studentTemplateList.toArray()));

    }

    /**
     * 将Sheet表中的数据封装对象中
     * @param workbook excel表中工作簿
     * @param clazz 需要返回对象的class
     * @param <T> 对象
     * @return list<T>
     */
    private static <T> List<T> populate(Workbook workbook,Class<T> clazz) throws IllegalAccessException, InstantiationException, ParseException {

        List<Map<Integer,Integer>> errorCell = new ArrayList<>();
        //主数据sheet
        Sheet dataSheet = workbook.getSheetAt(0);
        List<T> list = new ArrayList<>();
        //1.获取注解@Excel上的name属性
        Map<String, String> classData = getAnnotationExcelNameToMap(clazz);

        //获取sheet总行数
        int totalRows =dataSheet.getPhysicalNumberOfRows();
        //获取sheet总列数(数据一般从第二行开始)
        int totalColumn = dataSheet.getRow(2).getPhysicalNumberOfCells();

        for (int i = 2;i < totalRows; i++){
            if(dataSheet.getRow(i).getCell(0) == null || dataSheet.getRow(i).getCell(0).equals("")){
                break;
            }
            //构造业务对象
            T t  = clazz.newInstance();

            for(int j = 0; j< totalColumn; j++){
                //获取当前单元格
                Cell cell = dataSheet.getRow(i).getCell(j);
                //TODO 判断单元格是否为空，记录下来
                if(cell != null){
                    String cellValue  = getCellValueConvertString(cell);
                    if(cellValue != null){
                        System.out.println("当前单元格" + i + " ++ " + j);
                        String cellTitle = dataSheet.getRow(1).getCell(j).getStringCellValue();
                        classData.forEach((k, v)->{
                            //如果标题相等，则写入相对应的对象
                            if(v.equalsIgnoreCase(cellTitle)){
                                Field[] fieldsNewInstance = t.getClass().getDeclaredFields();
                                Arrays.stream(fieldsNewInstance).forEach(f->{
                                    if(f.getName().equalsIgnoreCase(k)){

                                        boolean flag = f.isAccessible();
                                        f.setAccessible(true);
                                        try {
                                            if(f.getType() == Map.class){
                                                //将对应关系取出后整理成Map
                                                Map<Integer,String> map = new HashMap<>();
                                                Sheet tempSheet = workbook.getSheet(k);
                                                sheetToMap(tempSheet).forEach((key,value)->{
                                                    if(value.equalsIgnoreCase(cellValue)){
                                                        map.put(key,value);
                                                    }
                                                });
                                                f.set(t,map);
                                            }else if(f.getType() == Boolean.class){
                                                f.set(t,Boolean.valueOf(cellValue));
                                            }else if(f.getType() == Date.class){
                                                System.out.println(cellValue);
                                                f.set(t,new SimpleDateFormat("yyyy-MM-dd").parse(cellValue));
                                            }
                                            else{
                                                f.set(t, cellValue);
                                            }
                                        } catch (ParseException | IllegalAccessException e) {
                                            e.printStackTrace();
                                        }
                                        f.setAccessible(flag);
                                    }
                                });
                            }
                        });
                    }

                }

            }
            list.add(t);
        }
        return list;
    }

    /**
     * 3.判断单元格里的值，用不同方式获取后转为String
     * @param cell 目标单元格
     * @return value 的String字符串
     */
    private static String getCellValueConvertString(Cell cell) {
        String cellValue;
        switch (cell.getCellType()){
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            case BOOLEAN:
                cellValue = String.valueOf(cell.getBooleanCellValue()).trim();
                break;
            case NUMERIC:
                //判断日期类型
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    SimpleDateFormat dateFormat=new SimpleDateFormat("yyyy-MM-dd");
                    cellValue = dateFormat.format(cell.getDateCellValue());
                }
                else{
                    cellValue = String.valueOf(cell.getNumericCellValue());
                }
                break;
            default:
                cellValue = "";
                break;
        }
        return cellValue;
    }

    /**
     * 1.获取注解@Excel上的name属性
     * @param clazz Class类
     * @param <T> Class类泛型
     * @return Map<String,String>
     */
    private static <T> Map<String, String> getAnnotationExcelNameToMap(Class<T> clazz) {
        //反射对象，获取注解Excel上的name
        Map<String,String> classData = new HashMap<>(4);
        Field[] fields = clazz.getDeclaredFields();
        Arrays.stream(fields).forEach(f->{
            f.setAccessible(true);
            classData.put(String.valueOf(f.getName()),f.getAnnotation(Excel.class).name());
        });
        return classData;
    }


    /**
     * 2.将隐藏表中下拉数据存入Map
     * @param sheet sheet表
     * @return Map<Integer,String>
     */
    private static Map<Integer,String> sheetToMap(Sheet sheet){
       Map<Integer,String> map = new HashMap<>(4);
       int totalRows =sheet.getPhysicalNumberOfRows();
       for (int i = 0; i < totalRows; i++) {
            map.put((int) sheet.getRow(i).getCell(0).getNumericCellValue(),sheet.getRow(i).getCell(1).getStringCellValue());
       }
       return map;
    }

}
