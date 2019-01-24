package com.yneusoft.euexcel.test.demo;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import com.yneusoft.euexcel.test.entity.report.StudentReport;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

/**
 * 下载带数据的excel表
 * @author 易用软件-朱新旺(zhuxinwang@aliyun.com)
 * @date 2019/1/21 0021 16:55
 */
public class Demo02 {

    static String PATH = "D:\\test\\workbook58.xls";

    public static void main(String[] args) throws IOException {

        long startTime = System.currentTimeMillis();

        List<StudentReport> studentReportList = new ArrayList<>();
        StudentReport studentReport = null;
        for(int i = 0 ;i < 1000; i++){

            studentReport = new StudentReport();
            studentReport.setName("朱新旺" + i);
            studentReport.setSex(true);
            studentReport.setIdType("身份证" + i);
            studentReport.setSchool("清华大学" + i);
            studentReport.setRegistrationDate(new Date());
            studentReport.setBirthday(new Date());
            studentReportList.add(studentReport);
        }


        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("主标题", "sheet表名"),
                StudentReport.class, studentReportList);

        FileOutputStream fileOut = new FileOutputStream(PATH);
        workbook.write(fileOut);

        System.out.println("运行总时间：" + (System.currentTimeMillis() - startTime));
    }


}
