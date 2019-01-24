package com.yneusoft.euexcel.test.entity.report;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;

import javax.validation.constraints.NotNull;
import java.io.Serializable;
import java.util.Date;
import java.util.Map;

/**
 * @author 易用软件-朱新旺(zhuxinwang@aliyun.com)
 * @date 2019/1/21 0021 17:48
 */
@Data
public class StudentReport implements Serializable {

    /**
     * 学生证件类型
     */
    @Excel(name = "证件类型")
    private String idType;

    /**
     * 学生姓名
     */
    @NotNull
    @Excel(name = "学生姓名")
    private String name;

    /**
     * 学生性别
     */
    @Excel(name = "学生性别")
    private Boolean sex;

    /**
     * 毕业学校
     */
    @Excel(name = "毕业学校")
    private String school;

    @Excel(name = "出生日期", databaseFormat = "yyyyMMddHHmmss", format = "yyyy-MM-dd")
    private Date birthday;

    @Excel(name = "进校日期", databaseFormat = "yyyyMMddHHmmss", format = "yyyy-MM-dd")
    private Date registrationDate;



}
