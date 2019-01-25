package com.yneusoft.euexcel.demo.service;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import com.yneusoft.euexcel.demo.entity.report.TeacherReport;
import com.yneusoft.euexcel.demo.entity.template.StudentTemplate;
import com.yneusoft.euexcel.entity.ResponseWrapper;
import com.yneusoft.euexcel.enums.ResultEnum;
import com.yneusoft.euexcel.exception.BusinessException;
import com.yneusoft.euexcel.tool.ExcelTool;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static com.yneusoft.euexcel.tool.ExcelTool.setCellBackground;

/**
 * excel服务层，调用Excel工具类
 * @author 易用软件-朱新旺(zhuxinwang@aliyun.com)
 * @date 2019/1/25 0025 17:40
 */
@Service
public class ExcelService {

    /**
     * 1.下载模版的方法
     * @param response HttpServletResponse 返回
     */
    public void templateDownload(HttpServletResponse response){

        String title = "这是主标题";
        String sheetName = "这是Sheet名称";

        StudentTemplate studentTemplate = new StudentTemplate();

        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams(title, sheetName),
                StudentTemplate.class, new ArrayList<>());

        //region 1.下拉数据的填充，这里可从存储过程中获取后赋值到Map<String,String>
        Map<String,String> idTypeMap = new HashMap<>();
        idTypeMap.put("1","身份证");
        idTypeMap.put("2","外国通行证");
        studentTemplate.setIdType(idTypeMap);

        Map<String,String> schoolMap = new HashMap<>();
        schoolMap.put("1","云南大学");
        schoolMap.put("2","外国语学院");
        studentTemplate.setSchool(schoolMap);
        //endregion

        try {
            ExcelTool.downExcel(response,workbook,title);
        } catch (IOException e) {
            throw new BusinessException(ResultEnum.EXCEL_TEMPLATE_ERROR);
        }

    }

    /**
     * 2.导出数据下载，导出报表
     * @param response HttpServletResponse 返回
     */
    public void reportDownload(HttpServletResponse response){

        String title = "这是主标题";
        String sheetName = "这是Sheet名称";

        //region 1.此处伪造从数据库中请求出来的数据
        List<TeacherReport> teacherReportList = new ArrayList<>();
        for(int i = 0 ;i < 1000; i++){
            TeacherReport teacherReport = new TeacherReport();
            teacherReport.setName("朱新旺" + i);
            teacherReport.setIdType("身份证" + i*i);
            teacherReport.setSchool("清华" + i +"大学");
            teacherReportList.add(teacherReport);
        }
        //endregion

        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams(title, sheetName),
                TeacherReport.class, teacherReportList);

        try {
            ExcelTool.downExcel(response,workbook,title);
        } catch (IOException e) {
            throw new BusinessException(ResultEnum.EXCEL_REPORT_DATA_ERROR);
        }
    }

    /**
     * 3.导入模版数据
     * @param multipartFile 请求的Excel文件
     * @return 导入结果
     */
    public ResponseWrapper importExcel(MultipartFile multipartFile){

        List<Map<Integer,Integer>> errorCell = new ArrayList<>();
        List<StudentTemplate> studentTemplateList = null;
        Workbook workbook = null;
        try {
            InputStream is = ExcelTool.uploadExcel(multipartFile);
            workbook = new HSSFWorkbook(is);
            studentTemplateList = ExcelTool.populateSheetToObj(workbook,StudentTemplate.class,2,errorCell);
        } catch (IOException | IllegalAccessException | ParseException | InstantiationException e) {
            throw new BusinessException(ResultEnum.EXCEL_IMPORT_DATA_ERROR);
        }

        //有错误写出文件
        if(errorCell.size() > 0){
            setCellBackground(workbook,errorCell);
            OutputStream os = null;
            try {
                workbook.write(os);
            } catch (IOException e) {
                e.printStackTrace();
            }

        }
        return ResponseWrapper.markSuccess(studentTemplateList);
    }

}
