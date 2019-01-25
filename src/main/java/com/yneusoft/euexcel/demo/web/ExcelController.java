package com.yneusoft.euexcel.demo.web;

import com.yneusoft.euexcel.demo.service.ExcelService;
import com.yneusoft.euexcel.entity.ResponseWrapper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;

/**
 * Excel控制器
 * @author 易用软件-朱新旺(zhuxinwang@aliyun.com)
 * @date 2019/1/25 0025 17:40
 */
@Controller
@RequestMapping(value = "excel")
public class ExcelController {

    @Autowired
    ExcelService excelService;

    @RequestMapping(value = "templateDownload")
    public void templateDownload(HttpServletResponse response){
        excelService.templateDownload(response);
    }

    @RequestMapping(value = "reportDownload")
    public void reportDownload(HttpServletResponse response){
        excelService.reportDownload(response);
    }

    @RequestMapping(value = "importExcel")
    @ResponseBody
    public ResponseWrapper importExcel(MultipartFile file){
        return excelService.importExcel(file);
    }

}
