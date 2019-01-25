package com.yneusoft.euexcel.enums;

import lombok.Getter;

/**
 * 接口返回码和返回值的枚举
 * @author 易用软件-朱新旺(zhuxinwang@aliyun.com)
 * @date 2018/11/13 0013 20:55
 */
@Getter
public enum ResultEnum {
    /**
     * 枚举类型
     */
    SUCCESS(10000,"成功"),
    FAILED(10001,"失败"),
    REQUEST_NOT_SUPPORTED(20000,"不支持的请求方式"),
    API_NOT_EXISTS(20001, "请求的接口不存在"),
    API_NOT_PER(20002, "没有该接口的访问权限"),
    PARAMS_ERROR(20003, "参数为空或格式错误"),
    SIGN_ERROR(20004, "数据签名错误"),
    API_DISABLE(20005, "查询权限已被限制"),
    UNKNOWN_IP(20006, "非法IP请求"),
    PARAMS_TYPE_ERROR(20007, "传输的参数类型有错误"),
    PARAMS_VALIDATION_ERROR(20008, "参数校验异常，请您检查填入的值"),
    PAGE_PATH_NOT_PER(20009, "您没有权限访问该界面的权限，请联系超级管理员"),
    DATABASE_ERROR(30000,"数据库异常"),
    DATABASE_PRO_ERROR(30001,"数据库异常<存储过程>"),
    AUTH_HEADER_NULL(40000,"携带参数的Token为空"),
    TOKEN_NOT_EXISTS(40001,"Token不存在"),
    TOKEN_ERROR(40002,"Token错误"),
    TOKEN_TIMEOUT(40003,"Token已失效"),
    TOKEN_SIGN_ERROR(40004,"Token签名错误"),
    CAPTCHA_ERROR(50000,"获取图片验证码错误"),
    CAPTCHA_TIMEOUT(50001,"图片验证码过期或不存在"),
    CAPTCHA_CHECK_ERROR(50002,"图片验证码校验错误"),
    WX_CODE_ERROR(50003,"微信code错误"),
    LOGIN_ERROR(60000,"用户名或密码错误"),
    EXIST_ADMIN_USER(60001,"该管理员账户已存在"),
    DISABLE_ADMIN_USER(60002,"该管理员账户已被禁用，请联系超级管理员"),
    OPENID_NOT_EXISTS(60003,"该openid未绑定用户，请跳转至登录界面登录"),
    USER_EXISTS(60004,"用户已存在，请跳转至登录界面登录"),
    USER_NOT_EXISTS(60005,"用户不存在，请跳转至注册界面注册"),
    COMPANY_NOT_EXISTS(60006,"公司绑定码不存在，请检查绑定码"),
    FILE_MAX_SIZE(60007,"上传的文件超过最大值"),
    COMPAY_ALREADY_BOUND(60008,"该绑定码对应公司当前已绑定"),
    COMPAY_BIND_SUCCESS(60009,"公司绑定成功"),
    MODEL_NOT_EXIST(70000,"短信模版不存在"),
    MESSAGE_NOT_EXIST(70001,"验证短信不存在"),
    MESSAGE_CODE_ERROR(70002,"短信验证码不存在或错误"),
    MESSAGE_SEND_ERROR(70003,"短信发送失败"),
    NOT_EXIST_CALLER(70004,"不存在的调用者"),
    MESSAGE_SEND_SUCCESS(70005,"短信发送成功"),
    CANNOT_COMMENT(80000, "该文章不可评论"),
    EXCEL_TEMPLATE_ERROR(80001, "Excel模版生成失败"),
    EXCEL_REPORT_DATA_ERROR(80002, "Excel导出数据报表失败"),
    EXCEL_IMPORT_DATA_ERROR(80003, "Excel导入数据失败"),
    EXCEL_DATA_CHECK_ERROR(80004, "Excel导入数据校验失败，详细错误请查看错误Excel"),
    HTTP_ERROR(90001,"http请求异常"),
    SYSTEM_ERROR(99999, "系统异常");

    private Integer code;
    private String message;

    ResultEnum(Integer code,String message){
        this.code = code;
        this.message = message;
    }
}
