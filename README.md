## EuExcel使用操作指南 
### 一、工具介绍
该工具类结合**POI** + **EasyPOI** ,充分使用EasyPOI的优点，使用类来做为Excel的模版，充分利用注解做校验。导入数据时，由于EasyPOI功能有限，所以配合POI使用，其中大部分类似在写EasyPOI的实现(主要基于POI做二次开发)，所以有些地方不够完善。如需添加和修改功能，可自行添加，或者联系<u>**zhuxinwang@aliyun.com**</u>进行添加后发布

>**当前版本 v1.0.0**

| 版本 | 日期 | 描述 |
| --- | --- | --- |
| v1.0.0 | 2019-06-26 | 初始上传，java Excel工具类，使用POI+EasyPOI |

### 二、工具结构
> 1.使用maven进行管理，需要引入pom文件的第三方jar包
```xml
<!-- fastjson 
https://mvnrepository.com/artifact/com.alibaba/fastjson -->
<dependency>
    <groupId>com.alibaba</groupId>
    <artifactId>fastjson</artifactId>
    <version>1.2.47</version>
</dependency>

<!-- 1.1  easypoi-base 导入导出的工具包,可以完成Excel导出,导入,Word的导出,Excel的导出功能 -->
<dependency>
    <groupId>cn.afterturn</groupId>    
    <artifactId>easypoi-base</artifactId>    
    <version>4.0.0</version>
</dependency>

<!-- 1.2  easypoi-annotation 基础注解包,作用与实体对象上,拆分后方便maven多工程的依赖管理 -->
<dependency>    
    <groupId>cn.afterturn</groupId>   
    <artifactId>easypoi-annotation</artifactId>   
    <version>4.0.0</version>
</dependency>

<!-- 2.1  Apache POI java操作excel https://mvnrepository.com/artifact/org.apache.poi/poi -->
<dependency>    
    <groupId>org.apache.poi</groupId>    
    <artifactId>poi</artifactId>    
    <version>4.0.1</version>
</dependency>
<!-- 2.2  Apache POI java操作excel https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml -->
<dependency>    
    <groupId>org.apache.poi</groupId>    
    <artifactId>poi-ooxml</artifactId>    
    <version>4.0.1</version>
</dependency>
```
>2.结构（最为主要的是在tool包下面的 **ExcelTool.java**）
````
EuExcel
├── .idea
├── .mvn
├── src
    ├── main
        ├── java
            ├── com.yneusoft.euexcel
                ├── demo
                    ├── entity
                    ├── service
                    ├── web
                ├── entity
                ├── enums
                ├── exception
                ├── handle
                ├── tool
        ├── resources
    ├── .......
````
##### 在上述结构中我们需要注意的只有两个地方，第一个是 **tool** 包（改包里具有工具类**ExcelTool.java**），其次需要注意的是com.yneufost.euexcel下的 **demo** 包（该包里**service**包里为主要调用ExcelTool.java的实现，请使用者参考编写代码）
>3.工具具体使用（真正使用的地方为   ---**region 实际调用** ---  注释的地方，其他的为构造虚拟数据）
---
>>3.1【示例】模版下载 templateDownload
````java
/** 
* 1.下载模版的方法 
* @param response HttpServletResponse 返回 
**/
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
````

3.1.1 调用ExcelTool中**templateDownload**的方法，该方法实例的使用需要先创建 **StudentTemplate** 类对象
3.1.2 需要填写 **excelName** 文件名称
3.1.3 需要填写 **title** 主题
3.1.4需要填写 **sheetName** Sheet表名
3.1.5 字段类型为 **Map** 的填充下拉数据
3.1.6 方法需要传入 **HTTPServletResponse** 提供下载

模版建类实例
````java
@Data
public class StudentTemplate implements Serializable {   
    /**     
    * 学生证件类型     
    */   
    @Excel(name = "证件类型")
    @NotNull    
    private Map<String,String> idType;   
    /**     
    * 学生姓名     
    */    
    @Excel(name = "学生姓名")    
    private String name;   
    /**     
    * 学生性别     
    */    
    @Excel(name = "学生性别")    
    @NotNull    
    private Boolean sex;   
    /**     
    * 毕业学校     
    */    
    @Excel(name = "毕业学校")    
    @NotNull    
    private Map<String,String> school;    
    @Excel(name = "出生日期")    
    @NotNull    
    private Date birthday;    
    @Excel(name = "进校日期")    
    @NotNull    
    private Date registrationDate;
}
````
3.1.7 @Excel 该注解中具有多个属性，其中使用 **name** 必须填写，这是生成Excel文件后的表头标题，还有其他属性，请参考EasyPOI文档 [http://easypoi.mydoc.io/](http://easypoi.mydoc.io/)
3.1.8 如单元格为下拉类型，请使用 **Map<String,String>** 作为字段类型
3.1.9 上传使用的是**同一模版处理，所以，可以使用注解校验，目前支持的注解有** @NotNull、@Size、@Max、@Min、@Pattern 

---
>>3.2 导出数据下载，导出报表reportDownload
````java
/** 
* 2.导出数据下载，导出报表 
* @param response HttpServletResponse 返回 
**/
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
````
3.2.1 调用ExcelTool中**reportDownload**的方法，该方法实例的使用需要先创建 **TeacherReport** 类对象
3.2.2 需要填写 **excelName** 文件名称
3.2.3 需要填写 **title** 主题
3.2.4需要填写 **sheetName** Sheet表名
3.2.5 方法需要传入 **HTTPServletResponse** 提供下载
````java
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
````
3.2.8@Excel 该注解中具有多个属性，其中使用 **name** 必须填写，这是生成Excel文件后的表头标题，还有其他属性，请参考EasyPOI

---
>>3.3导入模版数据
````java
/** 
* 3.导入模版数据 
* @param multipartFile 请求的Excel文件 
* @return 导入结果 
**/
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
````
3.3.1 调用ExcelTool中 **uploadExcelToWorkBook**  、**populateSheetToObj** 、**errorCellExcelFile** 的方法，
3.3.2 需要填写 **errorFileName** 如果有错误，导出错误的Excel的名称 
3.3.3 需要填写 **errorMap、errorCell** 主题
3.3.4需要填写 **studentTemplateList** 数据均正确时返回的信息   
#### 3.3中有两种返回方式，1.【失败】带错误文件的名称以及 Base64 的内容 2.【成功】正确的Json形式数据

### 三、注意
>1 3.1和3.2为请求直接下载
>2 导入数据可以在接口请求工具中查看请求结果，项目可直接运行，失败是的base64使用js将其还原成excel文件
>3【数据库多层json处理】由于成功的时候返回的多层json 数据，这里提供 Sql Server中解析多层 Json的示例
````SQL
DECLARE @json NVARCHAR(MAX)
SET @json =  N'[  
       { "id" : 2,"info": { "name": "John", "surname": "Smith" }, "age": 25 },  
       { "id" : 5,"info": { "name": "Jane", "surname": "Smith" }, "dob": "2005-11-04T12:00:00" }   ]'  
       
SELECT *  
FROM OPENJSON(@json)  
  WITH (id int 'strict $.id',  
        firstName nvarchar(50) '$.info.name',
        lastName nvarchar(50) $.info.surname',  
        age int, 
        dateOfBirth datetime2 '$.dob')
````
执行结果
| id | firstName | lastName | age | dateOfBirth |
| --- | --- | --- | --- | --- |
| 2 | John | Smith | 25 | NULL |
| 5 | Jane | Smith | NULL | 2005-11-04 12:00:00.0000000 |


>4 返回数据格式

失败
````JSON
{
    "success": false,
    "code": 80004,
    "message": "Excel导入数据校验失败，详细错误请查看错误Excel",-
    "data": {
        "name": "导入数据错误信息",
        "errorFileBase64": "data:application/vnd.ms-excel;base64,0M8R4KGxGuEAAAAAAAAAAAA......
    }
}   
````

成功
````JSON
{
    "success": true,
    "code": 10000,
    "message": "成功",-
    "data": [
    {
        "idType": {
            "name": "外国通行证",
            "id": "2"
        },
        "name": "朱新旺",
        "sex": true,
        "school": {
            "name": "云南大学",
            "id": "1"
        },
        "birthday": "2018-01-31T16:00:00.000+0000",
        "registrationDate": "2018-04-04T16:00:00.000+0000"
    }]
}
