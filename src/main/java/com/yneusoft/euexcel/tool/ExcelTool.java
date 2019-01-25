package com.yneusoft.euexcel.tool;

import cn.afterturn.easypoi.excel.annotation.Excel;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import javax.validation.constraints.*;
import java.io.*;
import java.lang.reflect.Field;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Excel工具类
 * @author 易用软件-朱新旺(zhuxinwang@aliyun.com)
 * @date 2019/1/23 0023 9:58
 */
@Component
public class ExcelTool {

    /**
     * 锁定单元格密码
     */
    private static String PROTECT_SHEET_PASSWORD = "EU_SOFT";



    /**
     * 1.将下拉列表的数据写入隐藏表及设置当前下拉列
     * @param workbook 工作簿
     * @param t 下载传输的下拉对象类型
     */
    public static <T> void setDropDownSheet(Workbook workbook, T t) throws IllegalAccessException, InstantiationException {
        Class clazz = t.getClass();
        Field[] fields = clazz.getDeclaredFields();
        //判断类型为List的，将设置为隐藏表；将其数据取出后写入到sheet表中
        if (fields.length > 0){
            for (Field f:fields) {
                if(f.getType() == Map.class){
                    //将第一张sheet表为list下的字段设置为下拉
                    setDropDownData(workbook,f,t);
                }
            }
        }
    }


    /**
     * 2.将下拉数据设置到主sheet表中
     * @param workbook 工作簿
     * @param field 反射的地段
     */
    private static <T> void setDropDownData(Workbook workbook, Field field, T t) throws IllegalAccessException {
        Sheet tempSheet = null;
        int index = 0;
        tempSheet = workbook.getSheetAt(0);
        field.setAccessible(true);
        Map<String,String> dropDownDataMap = (Map<String,String>) field.get(t);
        List<String> dropDownDataList = new ArrayList<>();
        // 拼接key-vlaue
        dropDownDataMap.forEach((k,v)->{
            dropDownDataList.add(k + "-" + v);
        });
        // 加载下拉列表内容
        DVConstraint constraint = DVConstraint.createExplicitListConstraint(dropDownDataList.toArray(new String[0]));
        // 把下拉内容加载到对应的列上
        Field[] fields = t.getClass().getDeclaredFields();
        List<String> fieldNames = new ArrayList<>();
        Arrays.asList(fields).forEach(f -> {
            fieldNames.add(f.getName());
        });
        index = Arrays.binarySearch(fieldNames.toArray(),field.getName());
        index = Math.abs(index);
        // 设置数据有效性加载在哪个单元格上,四个参数分别是：起始行、终止行、起始列、终止列
        CellRangeAddressList regions = new CellRangeAddressList(2,10000, index, index);
        // 设置具体数据
        HSSFDataValidation dataValidationList = new HSSFDataValidation(regions, constraint);
        // 在对应单元格上提示从下拉列表选择值
        String title = field.getAnnotation(Excel.class).name();
        dataValidationList.createPromptBox("", "请从下拉列表中选择" + title);
        tempSheet.addValidationData(dataValidationList);
    }


    /**
     * 3.设置下拉数据到隐藏表中
     * @param workbook 工作簿
     * @param field 下拉数据，反射的字段
     */
    private static void setDropDownDataToSheetHidden(Workbook workbook, Field field, Object object) throws IllegalAccessException {
        Sheet tempSheet = null;
        Row row = null;
        Cell cell1 = null;
        Cell cell2 = null;
        int rowValue = 0;
        tempSheet= workbook.createSheet(field.getName());
        field.setAccessible(true);
        Map<Integer,String> dropDownDataMap = (Map<Integer,String>) field.get(object);

        for (Map.Entry<Integer,String> entry : dropDownDataMap.entrySet()) {
            row = tempSheet.createRow(rowValue);
            cell1 = row.createCell(0);
            cell1.setCellValue(entry.getKey());
            cell2 = row.createCell(1);
            cell2.setCellValue(entry.getValue());
            rowValue = rowValue + 1;
        }

        //锁定Sheet表
        tempSheet.protectSheet(PROTECT_SHEET_PASSWORD);

        //隐藏tempSheet
        workbook.setSheetHidden(workbook.getSheetIndex(field.getName()),true);
    }



    /**
     * 4.将Sheet表中的数据封装对象中
     * @param workbook excel表中工作簿
     * @param clazz 需要返回对象的class
     * @param <T> 对象
     * @param dataStartColumn 数据开始行
     * @return list<T>
     */
    public static <T> List<T> populateSheetToObj(Workbook workbook,Class<T> clazz,Integer dataStartColumn,List<Map<Integer,Integer>> errorCell) throws IllegalAccessException, InstantiationException, ParseException {
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
                //当前单元格的值
                String cellValue = getCellValueConvertString(cell);
                Map<Integer,Integer> errorCellMap = new HashMap<>(1);
                for (Map.Entry<String,String> entry:titleMap.entrySet()){
                    //如果标题相等，则写入相对应的对象
                    if(entry.getValue().equalsIgnoreCase(cellTitle)){
                        Class businessClazz = t.getClass();
                        Field[] fieldsNewInstance = businessClazz.getDeclaredFields();
                        for(Field field:fieldsNewInstance){
                            if(field.getName().equalsIgnoreCase(entry.getKey())){
                                boolean flag = field.isAccessible();
                                field.setAccessible(true);
                                // 设置对象的属性值的同时判断属性上是否有对应注解的规则满足，不满足则计入错误列表
                                if(!annotationRule(field,cellValue)){
                                    errorCellMap.put(i,j);
                                    errorCell.add(errorCellMap);
                                    continue;
                                }
                                if(field.getType() == Map.class){
                                    //将对应关系取出后整理成Map
                                    field.set(t,stringToMap(cellValue));
                                } else if(field.getType() == Boolean.class){
                                    field.set(t,Boolean.valueOf(cellValue));
                                } else if(field.getType() == Date.class){
                                    field.set(t,new SimpleDateFormat("yyyy-MM-dd").parse(cellValue));
                                } else{
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
        return list;
    }

    /**
     * 5.将指定单元格背景替换颜色
     * @param workbook 工作簿
     * @param cellList 指定单元格列表
     */
    public static void setCellBackground(Workbook workbook,List<Map<Integer,Integer>> cellList){
        Sheet sheet = workbook.getSheetAt(0);
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellList.forEach(m->{
            m.forEach((k,v)->{
                sheet.getRow(k).createCell(v).setCellStyle(cellStyle);
            });
        });
    }



    /**
     * 6.判断值是否满足Annotation上的规则，已支持注解 [NotNull，Size，Max，Min，Pattern]
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
        //Min注解
        else if (field.isAnnotationPresent(Min.class)){
            Min minAnno = field.getAnnotation(Min.class);
            if(Integer.valueOf(value) < minAnno.value()){
                flag = false;
            }
        }
        //Pattern注解，正则表达式
        else if(field.isAnnotationPresent(Pattern.class)){
            Pattern patternAnno = field.getAnnotation(Pattern.class);
            if(!java.util.regex.Pattern.matches(patternAnno.regexp(),value)){
                flag = false;
            }
        }

        return flag;
    }



    /**
     * 7.判断单元格里的值，用不同方式获取后转为String
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
     * 8.获取注解@Excel上的name属性
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
     * 9.将字符串按照规则截取出来后放入Map中
     * @param mapString 字符串  1-身份证
     * @return Map数组
     */
    private static Map<String,String> stringToMap(String mapString){
        Map<String,String> dataMap = new HashMap<>(1);
        if(mapString != null){
            String[] str = mapString.split("-");
            dataMap.put("id",str[0]);
            dataMap.put("name",str[1]);
        }
        return dataMap;
    }

    /**
     * 10.将隐藏表中下拉数据存入Map
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



    /**
     * 11.下载Excel文件
     * @param response Http返回
     * @param workbook 工作簿
     * @param excelName 文件的名称
     */
    public static void downExcel(HttpServletResponse response, Workbook workbook, String excelName) throws IOException {
        response.setContentType("octets/stream");
        response.addHeader("Content-Disposition", "attachment;filename="+new String(excelName.getBytes("gb2312"), "ISO8859-1" )+".xls");
        OutputStream os = response.getOutputStream();
        workbook.write(os);
        os.flush();
        os.close();
        workbook.close();
    }

    /**
     * 12.上传Excel文件
     * @param multipartFile 文件
     * @return InputSteam流
     */
    public static InputStream uploadExcel(MultipartFile multipartFile) throws IOException {
        //参数为空，直接返回空
        if(multipartFile == null) {
            return null;
        }
        File file = multipartFileToFile(multipartFile);
        return new FileInputStream(file);
    }

    /**
     * 13.MultipartFile转File
     * @param multipartFile MultipartFile
     * @return File
     */
    private static File multipartFileToFile(MultipartFile multipartFile) throws IOException {
        File file = new File(Objects.requireNonNull(multipartFile.getOriginalFilename()));
        FileOutputStream fileOutputStream = null;
        fileOutputStream = new FileOutputStream(file);
        fileOutputStream.write(multipartFile.getBytes());
        fileOutputStream.close();
        multipartFile.transferTo(file);
        return file;
    }

}
