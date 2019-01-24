package com.yneusoft.euexcel.test.demo;

import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.util.PoiPublicUtil;
import com.yneusoft.euexcel.test.entity.report.StudentReport;
import com.yneusoft.euexcel.test.entity.template.StudentTemplate;
import org.apache.commons.lang3.builder.ReflectionToStringBuilder;
import org.springframework.util.Assert;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.List;

/**
 * 3.上传模版数据
 * @author 易用软件-朱新旺(zhuxinwang@aliyun.com)
 * @date 2019/1/21 0021 17:01
 */
public class Demo03 {

    static String PATH = "D:\\test\\workbook56.xls";

    public static void main(String[] args) {



        ImportParams params = new ImportParams();
        params.setNeedVerify(true);
        params.setTitleRows(1);
        params.setHeadRows(1);
        long start = System.currentTimeMillis();
        List<StudentReport> list = ExcelImportUtil.importExcel(
                new File(PATH),
                StudentReport.class, params);
        System.out.println(System.currentTimeMillis() - start);
        System.out.println(list.size());
        System.out.println(Arrays.toString(list.toArray()));
    }
}
