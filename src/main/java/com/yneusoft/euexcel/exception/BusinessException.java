package com.yneusoft.euexcel.exception;

import com.yneusoft.euexcel.enums.ResultEnum;
import lombok.Data;

/**
 * 自定义业务异常
 * @author 易用软件-朱新旺(zhuxinwang@aliyun.com)
 * @date 2018/11/13 0013 20:56
 */
@Data
public class BusinessException extends RuntimeException{
    /**
     * 错误码
     */
    private Integer code;

    public BusinessException() {}

    public BusinessException(ResultEnum resultEnum) {
        super(resultEnum.getMessage());
        this.code = resultEnum.getCode();
    }
}
