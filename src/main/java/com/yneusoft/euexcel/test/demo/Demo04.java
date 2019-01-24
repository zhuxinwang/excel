package com.yneusoft.euexcel.test.demo;

import cn.afterturn.easypoi.excel.annotation.Excel;
import com.yneusoft.euexcel.test.entity.template.StudentTemplate;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.text.DecimalFormat;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Map导入
 * @author 易用软件-朱新旺(zhuxinwang@aliyun.com)
 * @date 2019/1/23 0023 10:02
 */
public class Demo04 {


    static String PATH = "D:\\test\\workbook56.xls";

    public static void main(String[] args) throws IOException, IllegalAccessException, InstantiationException {
        List<StudentTemplate> studentTemplateList = null;


        InputStream is = new FileInputStream(PATH);
        Workbook workbook = new HSSFWorkbook(is);
        Sheet dataSheet = workbook.getSheetAt(0);
        studentTemplateList = populate(dataSheet,StudentTemplate.class);

    }

    /**
     * 将Sheet表中的数据封装对象中
     * @param dataSheet excel表中sheet表
     * @param clazz 需要返回对象的class
     * @param <T> 对象
     * @return list<T>
     */
    private static <T> List<T> populate(Sheet dataSheet,Class<T> clazz) throws IllegalAccessException, InstantiationException {
        //反射对象，获取注解上的name
        Map<String,String> classData = new HashMap<>();
        Field[] fields = clazz.getDeclaredFields();
        Arrays.stream(fields).forEach(f->{
            Excel excelAnno = f.getAnnotation(Excel.class);
            f.setAccessible(true);
            classData.put(String.valueOf(f.getName()),excelAnno.name());
        });

        //region 1.打印反射出来的键值对对象
        classData.forEach((k,v) -> {
            System.out.println(k + "   + +   " + v);
        });
        //endregion

        //获取sheet总行数
        int totalRows =dataSheet.getPhysicalNumberOfRows();
        //获取sheet总列数(数据一般从第二行开始)
        int totalColumn = dataSheet.getRow(2).getPhysicalNumberOfCells();
        for (int i = 2;i < totalRows; i++){
            //构造业务对象
            T t  = clazz.newInstance();
            for(int j = 0; j< totalColumn; j++){
                Cell cell = dataSheet.getRow(i).getCell(j);

                //TODO 判断单元格是否为空，记录下来

                String cellValue = "";

                switch (cell.getCellType()){
                    case STRING:
                        cellValue = cell.getStringCellValue();
                        break;
                    case BOOLEAN:
                        cellValue = String.valueOf(cell.getBooleanCellValue()).trim();
                        break;
                    case NUMERIC:
                        //判断日期类型
                        if
                        cellValue = String.valueOf(cell.getNumericCellValue());
                        break;
                    default:
                        cellValue = "";
                        break;
                }

            //将单元格上的值赋值个对象，如果是Map，则查询隐藏表后封装T对象
            System.out.println(cellValue);
            }

        }

        return null;
    }

}
