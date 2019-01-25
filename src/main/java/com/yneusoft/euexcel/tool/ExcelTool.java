package com.yneusoft.euexcel.tool;

import cn.afterturn.easypoi.excel.annotation.Excel;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.springframework.stereotype.Component;

import java.lang.reflect.Field;
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
     * @return 新的工作簿
     */
    public static <T> Workbook setDropDownSheet(Workbook workbook, T t) throws IllegalAccessException, InstantiationException {
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
        return workbook;
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
        Map<Integer,String> dropDownDataMap = (Map<Integer,String>) field.get(t);
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


}
