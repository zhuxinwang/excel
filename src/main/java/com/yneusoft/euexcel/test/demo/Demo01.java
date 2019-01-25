package com.yneusoft.euexcel.test.demo;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import com.yneusoft.euexcel.test.entity.template.StudentTemplate;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import static com.yneusoft.euexcel.tool.ExcelTool.setDropDownSheet;


/**
 * 1.下载指定模版
 * @author 易用软件-朱新旺(zhuxinwang@aliyun.com)
 * @date 2019/1/24 0021 16:47
 */
public class Demo01 {

    static String PATH = "D:\\test\\workbook66.xls";

    public static void main(String[] args) throws IOException {

        long startTime = System.currentTimeMillis();

        StudentTemplate studentTemplate = new StudentTemplate();
        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("主标题", "sheet表名"),
                StudentTemplate.class, new ArrayList<>());

        Map<Integer,String> idTypeMap = new HashMap<>();
        idTypeMap.put(1,"身份证");
        idTypeMap.put(2,"外国通行证");
        studentTemplate.setIdType(idTypeMap);

        Map<Integer,String> schoolMap = new HashMap<>();
        schoolMap.put(1,"云南大学");
        schoolMap.put(2,"外国语学院");
        studentTemplate.setSchool(schoolMap);


        try {
            workbook = setDropDownSheet(workbook, studentTemplate);
        } catch (IllegalAccessException | InstantiationException e) {
            e.printStackTrace();
        }

        FileOutputStream fileOut = new FileOutputStream(PATH);
        workbook.write(fileOut);

        System.out.println("运行总时间：" + (System.currentTimeMillis() - startTime));
    }
}
