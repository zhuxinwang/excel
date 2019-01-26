package com.yneusoft.euexcel.demo.service;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import com.yneusoft.euexcel.demo.entity.report.TeacherReport;
import com.yneusoft.euexcel.demo.entity.template.StudentTemplate;
import com.yneusoft.euexcel.entity.ResponseWrapper;
import com.yneusoft.euexcel.enums.ResultEnum;
import com.yneusoft.euexcel.exception.BusinessException;
import com.yneusoft.euexcel.tool.ExcelTool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * excel服务层，调用Excel工具类
 * @author 易用软件-朱新旺(zhuxinwang@aliyun.com)
 * @date 2019/1/25 0025 17:40
 */
@Slf4j
@Service
public class ExcelService {

    /**
     * 1.下载模版的方法
     * @param response HttpServletResponse 返回
     */
    public void templateDownload(HttpServletResponse response){
        String excelName = "这是文件名称";
        String title = "这是主标题";
        String sheetName = "这是Sheet名称";

        //region 1.模拟数据，下拉数据的填充，这里可从存储过程中获取后赋值到Map<String,String>
        StudentTemplate studentTemplate = new StudentTemplate();
        Map<String,String> idTypeMap = new HashMap<>();
        idTypeMap.put("1","身份证");
        idTypeMap.put("2","外国通行证");
        studentTemplate.setIdType(idTypeMap);

        Map<String,String> schoolMap = new HashMap<>();
        schoolMap.put("1","云南大学");
        schoolMap.put("2","外国语学院");
        studentTemplate.setSchool(schoolMap);
        //endregion

        //region 实际调用
        try {
            ExcelTool.templateDownload(response,excelName,title,sheetName,studentTemplate);
        } catch (IOException | IllegalAccessException | InstantiationException e) {
            throw new BusinessException(ResultEnum.EXCEL_TEMPLATE_ERROR);
        }
        //endregion

    }

    /**
     * 2.导出数据下载，导出报表
     * @param response HttpServletResponse 返回
     */
    public void reportDownload(HttpServletResponse response){

        String excelName = "这是文件名称";
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

        //region 实际调用
        try {
            ExcelTool.reportDownload(response,excelName,title,sheetName,teacherReportList);
        } catch (IOException e) {
            throw new BusinessException(ResultEnum.EXCEL_REPORT_DATA_ERROR);
        }
        //endregion
    }

    /**
     * 3.导入模版数据
     * @param multipartFile 请求的Excel文件
     * @return 导入结果
     */
    public ResponseWrapper importExcel(MultipartFile multipartFile) {
        //如果有错误，导出错误的Excel的名称
        String errorFileName = "导入数据错误信息";
        //存放错误信息的Map及错误单元格的errorCell
        Map<String,Object> errorMap = null;
        List<Map<Integer, Integer>> errorCell = new ArrayList<>();
        //数据均正确时返回的信息
        List<StudentTemplate> studentTemplateList = null;


        //region 实际调用
        Workbook workbook = null;
        try {
            //导入文件获取成Workbook
            workbook = ExcelTool.uploadExcelToWorkBook(multipartFile);
            //将excel表中数据导入list中
            studentTemplateList = ExcelTool.populateSheetToObj(workbook, StudentTemplate.class, 2, errorCell);
            //将错误信息记录Base64
            errorMap = ExcelTool.errorCellExcelFile(workbook,errorCell,errorFileName);
            if(errorMap != null){
                return ResponseWrapper.markCustom(false,ResultEnum.EXCEL_DATA_CHECK_ERROR,errorMap);
            }else{

                // TODO 在这里可以将正确的数据进行自己的逻辑处理，如放入数据库等。

                return ResponseWrapper.markSuccess(studentTemplateList);
            }
        } catch (IOException | IllegalAccessException | ParseException | InstantiationException e) {
            log.info("导入数据异常，异常信息" + e.getMessage());
            throw new BusinessException(ResultEnum.EXCEL_IMPORT_DATA_ERROR);
        }
        //endregion
    }
}
