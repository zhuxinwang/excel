package com.yneusoft.euexcel.entity;

import com.yneusoft.euexcel.enums.ResultEnum;
import com.yneusoft.euexcel.exception.BusinessException;
import lombok.Data;

/**
 * api接口返回
 * @author 易用软件-朱新旺(zhuxinwang@aliyun.com)
 * @date 2018/11/13 0013 20:59
 */
@Data
public class ResponseWrapper {
    /**
     * 是否成功
     */
    private boolean success;
    /**
     * 返回码
     */
    private Integer code;
    /**
     * 返回信息
     */
    private String message;
    /**
     * 具体返回的实体类
     */
    private Object data;

    /**
     * 自定义返回结果
     * 建议使用统一的返回结果，特殊情况可以使用此方法
     * @param success 是否成功
     * @param resultEnum 返回枚举
     * @param data 具体返回对象
     * @return ResponseWrapper封装类
     */
    public static ResponseWrapper markCustom(boolean success, ResultEnum resultEnum, Object data){
        ResponseWrapper responseWrapper = new ResponseWrapper();
        responseWrapper.setSuccess(success);
        responseWrapper.setCode(resultEnum.getCode());
        responseWrapper.setMessage(resultEnum.getMessage());
        responseWrapper.setData(data);
        return responseWrapper;
    }


    /**
     * 自定义返回结果【错误时】只需要传输错误码和消息
     * @param resultEnum 枚举信息
     * @return ResponseWrapper封装类
     */
    public static ResponseWrapper markCustom(ResultEnum resultEnum){
        ResponseWrapper responseWrapper = new ResponseWrapper();
        responseWrapper.setSuccess(false);
        responseWrapper.setCode(resultEnum.getCode());
        responseWrapper.setMessage(resultEnum.getMessage());
        responseWrapper.setData(null);
        return responseWrapper;
    }

    /**
     * 自定义返回结果【业务错误】只需要传输错误码和消息
     * @param businessException 枚举信息
     * @return ResponseWrapper封装类
     */
    public static ResponseWrapper markCustom(BusinessException businessException){
        ResponseWrapper responseWrapper = new ResponseWrapper();
        responseWrapper.setSuccess(false);
        responseWrapper.setCode(businessException.getCode());
        responseWrapper.setMessage(businessException.getMessage());
        responseWrapper.setData(null);
        return responseWrapper;
    }


    /**
     * 返回错误
     * @return ResponseWrapper封装类
     */
    public static ResponseWrapper markError() {
        ResponseWrapper responseWrapper = new ResponseWrapper();
        responseWrapper.setSuccess(false);
        responseWrapper.setCode(ResultEnum.FAILED.getCode());
        responseWrapper.setMessage(ResultEnum.FAILED.getMessage());
        return responseWrapper;
    }

    /**
     * 自定义返回错误信息
     * @param resultEnum 返回枚举
     * @return 错误信息
     */
    public static ResponseWrapper markError(ResultEnum resultEnum) {
        ResponseWrapper responseWrapper = new ResponseWrapper();
        responseWrapper.setSuccess(false);
        responseWrapper.setCode(resultEnum.getCode());
        responseWrapper.setMessage(resultEnum.getMessage());
        return responseWrapper;
    }

    /**
     * 自定义返回错误信息
     * @param resultEnum 返回枚举
     * @param object 错误对象
     * @return 错误信息
     */
    public static ResponseWrapper markError(ResultEnum resultEnum, Object object) {
        ResponseWrapper responseWrapper = new ResponseWrapper();
        responseWrapper.setSuccess(false);
        responseWrapper.setCode(resultEnum.getCode());
        responseWrapper.setMessage(resultEnum.getMessage());
        responseWrapper.setData(object);
        return responseWrapper;
    }


    /**
     * 查询成功且有数据
     * @param data 需要返回的对象
     * @return ResponseWrapper封装类
     */
    public static ResponseWrapper markSuccess(Object data) {
        ResponseWrapper responseWrapper = new ResponseWrapper();
        responseWrapper.setSuccess(true);
        responseWrapper.setCode(ResultEnum.SUCCESS.getCode());
        responseWrapper.setMessage(ResultEnum.SUCCESS.getMessage());
        responseWrapper.setData(data);
        return responseWrapper;
    }


    /**
     * 查询成功且有数据
     * @param resultEnum 返回的错误类型
     * @return ResponseWrapper封装类
     */
    public static ResponseWrapper markSuccess(Boolean success, ResultEnum resultEnum) {
        ResponseWrapper responseWrapper = new ResponseWrapper();
        responseWrapper.setSuccess(success);
        responseWrapper.setCode(resultEnum.getCode());
        responseWrapper.setMessage(resultEnum.getMessage());
        responseWrapper.setData(null);
        return responseWrapper;
    }

    /**
     * 查询成功无数据
     * @return ResponseWrapper封装类
     */
    public static ResponseWrapper markSuccess() {
        ResponseWrapper responseWrapper = new ResponseWrapper();
        responseWrapper.setSuccess(true);
        responseWrapper.setCode(ResultEnum.SUCCESS.getCode());
        responseWrapper.setMessage(ResultEnum.SUCCESS.getMessage());
        responseWrapper.setData(null);
        return responseWrapper;
    }
}
