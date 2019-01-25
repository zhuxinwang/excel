package com.yneusoft.euexcel.handle;

import com.yneusoft.euexcel.entity.ResponseWrapper;
import com.yneusoft.euexcel.enums.ResultEnum;
import com.yneusoft.euexcel.exception.BusinessException;
import lombok.extern.slf4j.Slf4j;
import org.apache.tomcat.util.http.fileupload.FileUploadBase;
import org.springframework.http.converter.HttpMessageConversionException;
import org.springframework.validation.BindException;
import org.springframework.validation.BindingResult;
import org.springframework.validation.FieldError;
import org.springframework.web.HttpRequestMethodNotSupportedException;
import org.springframework.web.bind.MethodArgumentNotValidException;
import org.springframework.web.bind.MissingServletRequestParameterException;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestControllerAdvice;
import org.springframework.web.servlet.NoHandlerFoundException;

import javax.validation.ConstraintViolationException;
import java.sql.SQLException;
import java.util.HashMap;
import java.util.Map;

/**
 * 全局异常处理1
 * @author 易用软件-朱新旺(zhuxinwang@aliyun.com)
 * @date 2018/11/13 0013 20:58
 */
@Slf4j
@RestControllerAdvice
public class GlobalExceptionHandler {
    /**
     * 运行时异常捕获
     * @param e 异常
     * @return Json对象
     */
    @ResponseBody
    @ExceptionHandler(RuntimeException.class)
    public Object handlerRuntimeException(RuntimeException e) {
        // 参数校验异常
        if(e instanceof ConstraintViolationException){
            log.error("[参数校验异常]：" + e.getMessage());
            String[] errorMessage = e.getMessage().split(",");
            //统一返回异常样式，切割字符串
            Map<String,String> resultData = new HashMap<>(4);
            for (String item:errorMessage) {
                String[] message = item.split(":");
                String key = message[0].split("\\.")[1];
                resultData.put(key,message[1]);
            }
            return ResponseWrapper.markCustom(false, ResultEnum.PARAMS_VALIDATION_ERROR,resultData);
        }

        // http请求异常
        else if (e instanceof HttpMessageConversionException){
            log.error("[http请求异常]：" + e.getMessage());
            return ResponseWrapper.markCustom(ResultEnum.HTTP_ERROR);
        }


        // 业务异常
        else if(e instanceof BusinessException){
            log.error("[业务异常]：" + e.getMessage());
            BusinessException businessException = (BusinessException)e;
            return ResponseWrapper.markCustom(businessException);
        }

        else{
            //系统错误
            log.error("[未知异常]：" + e.getMessage(),e);
            return ResponseWrapper.markCustom(ResultEnum.SYSTEM_ERROR);
        }
    }

    /**
     * 全局异常捕捉处理
     * @param e 异常
     * @return Json对象
     */
    @ResponseBody
    @ExceptionHandler(Exception.class)
    public Object handlerException(Exception e){
        // 不支持的请求方式
        if(e instanceof HttpRequestMethodNotSupportedException){
            log.error("[不支持的请求方式]：" + (((HttpRequestMethodNotSupportedException) e).getMethod()));
            return ResponseWrapper.markCustom(ResultEnum.REQUEST_NOT_SUPPORTED);
        }
        // 访问接口不存在
        else if(e instanceof NoHandlerFoundException){
            log.error("[访问接口不存在]：" + ((NoHandlerFoundException) e).getRequestURL());
            return ResponseWrapper.markCustom(ResultEnum.API_NOT_EXISTS);
        }


        // 接口参数校验异常
        else if(e instanceof MethodArgumentNotValidException || e instanceof BindException){
            //确定异常类型后将错误信息取出
            BindingResult bindingResult;
            if (e instanceof MethodArgumentNotValidException){
                bindingResult = ((MethodArgumentNotValidException) e).getBindingResult();
            }else{
                bindingResult = ((BindException) e).getBindingResult();
            }
            StringBuilder errorMessage = new StringBuilder();
            Map<String,Object> validParam = new HashMap<>(8);
            for (FieldError fieldError : bindingResult.getFieldErrors()) {
                validParam.put(fieldError.getField(),fieldError.getDefaultMessage());
                errorMessage.append(fieldError.getField()).append(":").append(fieldError.getDefaultMessage()).append(",");
            }
            log.error("[Controller参数校验异常]：" + errorMessage);
            return ResponseWrapper.markCustom(false,ResultEnum.PARAMS_VALIDATION_ERROR,validParam);
        }


        // 缺少参数异常
        else if(e instanceof MissingServletRequestParameterException){
            log.error("[缺少参数异常]：" + e.getMessage());
            return ResponseWrapper.markCustom(ResultEnum.PARAMS_ERROR);
        }

        // 请求参数类型不匹配
        else if(e instanceof IllegalArgumentException){
            log.error("[请求参数类型不匹配]：" + e.getMessage());
            return ResponseWrapper.markCustom(ResultEnum.PARAMS_ERROR);
        }

        // 数据库异常
        else if(e instanceof SQLException){
            log.error("[数据库异常]：" + e.getMessage());
            return ResponseWrapper.markCustom(ResultEnum.DATABASE_ERROR);
        }

        // 文件size
        else if(e instanceof FileUploadBase.FileSizeLimitExceededException){
            log.error("[文件上传]：" + e.getMessage());
            return ResponseWrapper.markCustom(ResultEnum.FILE_MAX_SIZE);
        }

        else{
            log.error("[未知异常：]" + e.getMessage(),e);
            //系统错误
            return ResponseWrapper.markCustom(ResultEnum.SYSTEM_ERROR);
        }

    }
}
