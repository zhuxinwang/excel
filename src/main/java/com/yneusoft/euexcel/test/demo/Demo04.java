package com.yneusoft.euexcel.test.demo;

import cn.afterturn.easypoi.excel.annotation.Excel;
import com.yneusoft.euexcel.test.entity.template.StudentTemplate;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import javax.validation.constraints.*;
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


    static String PATH = "D:\\test\\workbook66.xls";

    public static void main(String[] args) throws IOException, IllegalAccessException, InstantiationException, ParseException {

        long startTime = System.currentTimeMillis();

        List<StudentTemplate> studentTemplateList = null;


        InputStream is = new FileInputStream(PATH);
        Workbook workbook = new HSSFWorkbook(is);
        studentTemplateList = populate(workbook,StudentTemplate.class,2);

        System.out.println("运行总时间：" + (System.currentTimeMillis() - startTime));

        System.out.println(Arrays.toString(studentTemplateList.toArray()));

    }

    /**
     * 将Sheet表中的数据封装对象中
     * @param workbook excel表中工作簿
     * @param clazz 需要返回对象的class
     * @param <T> 对象
     * @param dataStartColumn 数据开始行
     * @return list<T>
     */
    private static <T> List<T> populate(Workbook workbook,Class<T> clazz,Integer dataStartColumn) throws IllegalAccessException, InstantiationException, ParseException {

        List<Map<Integer,Integer>> errorCell = new ArrayList<>();
        //主数据sheet
        Sheet dataSheet = workbook.getSheetAt(0);
        List<T> list = new ArrayList<>();
        //1.获取注解@Excel上的name属性
        Map<String, String> titleMap = getAnnotationExcelNameToMap(clazz);

        //获取sheet总行数
        int totalRows =dataSheet.getPhysicalNumberOfRows();
        //获取sheet总列数
        int totalColumn = dataSheet.getRow(dataStartColumn).getPhysicalNumberOfCells();
        for (int i = dataStartColumn ;i < totalRows; i++){
            //构造业务对象
            T t  = clazz.newInstance();

            for(int j = 0; j< totalColumn; j++){

                //获取当前单元格
                Cell cell = dataSheet.getRow(i).getCell(j);

                //当前单元格的标题title
                String cellTitle = dataSheet.getRow(dataStartColumn-1).getCell(j).getStringCellValue();

                String cellValue = getCellValueConvertString(cell);

                Map<Integer,Integer> errorCellMap = new HashMap<>(1);

                System.out.println("当前Cell 的值： "+ cellValue);
                System.out.println("当前单元格" + i + " ++ " + j);
                for (Map.Entry<String,String> entry:titleMap.entrySet()){
                    //如果标题相等，则写入相对应的对象
                    if(entry.getValue().equalsIgnoreCase(cellTitle)){
                        Class businessClazz = t.getClass();
                        Field[] fieldsNewInstance = businessClazz.getDeclaredFields();
                        for(Field field:fieldsNewInstance){
                            if(field.getName().equalsIgnoreCase(entry.getKey())){
                                boolean flag = field.isAccessible();
                                field.setAccessible(true);
                                //TODO 设置对象的属性值的同时判断属性上是否有对应注解的规则满足，不满足则计入错误列表
                                if(!annotationRule(field,cellValue)){
                                    errorCellMap.put(i,j);
                                    errorCell.add(errorCellMap);
                                    continue;
                                }
                                if(field.getType() == Map.class){
                                    //将对应关系取出后整理成Map
                                    field.set(t,stringToMap(cellValue));
                                }else if(field.getType() == Boolean.class){
                                    field.set(t,Boolean.valueOf(cellValue));
                                }else if(field.getType() == Date.class){
                                    field.set(t,new SimpleDateFormat("yyyy-MM-dd").parse(cellValue));
                                }
                                else{
                                    field.set(t, cellValue);
                                }

                                field.setAccessible(flag);
                            }
                        }
                    }
                }

            }
            list.add(t);
        }

        System.out.println(Arrays.toString(errorCell.toArray()));
        return list;
    }

    /**
     * 4.判断值是否满足Annotation上的规则
     * @param field 反射字段
     * @param value 需要校验的值
     * @return true/false
     */
    public static Boolean annotationRule(Field field,String value){
        boolean flag = true;
        //NotNull注解
        if(field.isAnnotationPresent(NotNull.class)){
            if(value == null || "".equals(value)){
                flag = false;
            }
        }
        //Size注解
        else if(field.isAnnotationPresent(Size.class)){
            Size sizeAnno = field.getAnnotation(Size.class);
            if(sizeAnno.min() < Integer.valueOf(value) && Integer.valueOf(value) > sizeAnno.max()){
                flag = false;
            }
        }
        //Max注解
        else if(field.isAnnotationPresent(Max.class)){
            Max maxAnno = field.getAnnotation(Max.class);
            if(Integer.valueOf(value) > maxAnno.value()){
                flag = false;
            }
        }
        //Mix注解
        else if (field.isAnnotationPresent(Min.class)){
            Min minAnno = field.getAnnotation(Min.class);
            if(Integer.valueOf(value) < minAnno.value()){
                flag = false;
            }
        }
        //正则表达式
        else if(field.isAnnotationPresent(Pattern.class)){
            Pattern patternAnno = field.getAnnotation(Pattern.class);
            if(!java.util.regex.Pattern.matches(patternAnno.regexp(),value)){
                flag = false;
            }
        }

        return flag;
    }


    /**
     * 3.判断单元格里的值，用不同方式获取后转为String
     * @param cell 目标单元格
     * @return value 的String字符串
     */
    private static String getCellValueConvertString(Cell cell) {
        String cellValue = null;
        if(cell != null){
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
                    cellValue = null;
                    break;
            }
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
     * 3.将字符串按照规则截取出来后放入Map中
     * @param mapString 字符串  1-身份证
     * @return Map数组
     */
    public static Map<Integer,String> stringToMap(String mapString){
        Map<Integer,String> dateMap = new HashMap<>(1);
        if(mapString != null){
        String[] str = mapString.split("-");
            dateMap.put(Integer.valueOf(str[0]),str[1]);
        }
        return dateMap;
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
