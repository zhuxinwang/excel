package com.yneusoft.euexcel.test.demo;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import com.yneusoft.euexcel.test.entity.report.StudentReport;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * 下载带数据的excel表
 * @author 易用软件-朱新旺(zhuxinwang@aliyun.com)
 * @date 2019/1/21 0021 16:55
 */
public class Demo02 {

    static String PATH = "D:\\test\\workbook56.xls";

    public static void main(String[] args) throws IOException {
        StudentReport studentReport = new StudentReport();
        studentReport.setName("朱新旺");
        studentReport.setSex(true);
        studentReport.setIdType("身份证");
        studentReport.setSchool("清华大学");
        studentReport.setRegistrationDate(new Date());
        studentReport.setBirthday(new Date());
        StudentReport studentReport1 = new StudentReport();
        studentReport1.setName("朱大大");
        studentReport1.setSex(true);
        studentReport1.setIdType("港澳台通行证");
        studentReport1.setSchool("同济大学");
        studentReport1.setRegistrationDate(new Date());
        studentReport1.setBirthday(new Date());
        List<StudentReport> studentReportList = new ArrayList<>();
        studentReportList.add(studentReport);
        studentReportList.add(studentReport1);

        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("主标题", "sheet表名"),
                StudentReport.class, studentReportList);

        FileOutputStream fileOut = new FileOutputStream(PATH);
        workbook.write(fileOut);
    }


}
