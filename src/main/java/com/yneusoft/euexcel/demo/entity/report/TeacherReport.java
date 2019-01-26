package com.yneusoft.euexcel.demo.entity.report;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;

import javax.validation.constraints.NotNull;
import java.io.Serializable;
import java.util.Date;

/**
 * @author 易用软件-朱新旺(zhuxinwang@aliyun.com)
 * @date 2019/1/21 0021 17:48
 */
@Data
public class TeacherReport implements Serializable {
    /**
     * 教师姓名
     */
    @Excel(name = "教师姓名")
    private String name;
    /**
     * 教师证件类型
     */
    @Excel(name = "证件类型")
    private String idType;
    /**
     * 教师学校
     */
    @Excel(name = "教师学校")
    private String school;

}
