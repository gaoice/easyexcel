# easyexcel
## 概述

Excel快速生成工具

在现有实体类的基础上构建最简单的Excel，代码只需4行。



## 使用

`Maven`

```xml
<dependency>
    <groupId>com.gaoice</groupId>
    <artifactId>easyexcel</artifactId>
    <version>1.0</version>
</dependency>
```



## Starter

现在可以在 Spring Boot 中快速使用：[easyexcel-spring-boot-starter](https://github.com/gaoice/easyexcel-spring-boot-starter)



## 新版本

- v 1.1，支持 List<Map<?,?>> 类型，支持链式调用，详见 [MapListTests.java](https://github.com/gaoice/easyexcel/blob/master/src/test/java/com/gaoice/easyexcel/test/MapListTests.java) 。




## 示例

实体类：

`Student{name, idcard, sex, ...}`

生成SXSSFWorkbook：

```java
String sheetName = "sheet name";
String[] columnNames = {"姓名", "身份证号", "性别", ...};
String[] classFieldNames = {"name", "idcard", "sex", ...};
SXSSFWorkbook workbook = ExcelBuilder.createWorkbook(
    new SheetInfo(sheetName, columnNames, classFieldNames, studentList));
```

web下直接写入HttpServletResponse的OutputStream中：

```java
ExcelBuilder.writeOutputStream(
    new SheetInfo(sheetName, columnNames, classFieldNames, studentList), 
    response.getOutputStream());
```

如果sex字段是int类型的，我们可以为sex字段添加一个Converter（Lambda表达式）在构建Excel时转换为中文：

```java
sheetInfo.putConverter("sex", 
	(sheetInfo1, o, listIndex, columnIndex) -> o.equals(1) ? "男生" : "女生");
```

完整的使用方法示例详见 [Test文件](https://github.com/gaoice/easyexcel/blob/master/src/test/java/com/gaoice/easyexcel/test/ExcelBuilderTests.java) 。
