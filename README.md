# easyexcel
## 概述

采用极简设计的 Excel 生成和读取工具，几行代码即可把 `List` 集合生成为 Excel 或者从 Excel 读取为 `List` 集合。



## 使用

`Maven`

```xml
<dependency>
    <groupId>com.gaoice</groupId>
    <artifactId>easyexcel</artifactId>
    <version>1.1</version>
</dependency>
```



## Starter

现在可以在 Spring Boot 中快速使用：[easyexcel-spring-boot-starter](https://github.com/gaoice/easyexcel-spring-boot-starter) 



## 新版本

- v 2.0，支持 Excel 的读取，详见 [ExcelReaderTests.java](https://github.com/gaoice/easyexcel/blob/master/src/test/java/com/gaoice/easyexcel/test/reader/ExcelReaderTests.java) 




## 示例

实体类：

`Student{name, idcard, gender, ...}`

#### 写 Excel

生成 `SXSSFWorkbook`：

```java
String[] classFieldNames = {"name", "idcard", "gender", ...};
SXSSFWorkbook workbook = ExcelWriter.createWorkbook(new SheetInfo(classFieldNames, studentList));
```

web 下直接写入 `HttpServletResponse` 的 `OutputStream` 中：

```java
ExcelWriter.writeOutputStream(new SheetInfo(classFieldNames, studentList), response.getOutputStream());
```

如果 gender 字段是int类型的，我们可以为 gender 字段添加一个 `FieldValueConverter`（Lambda表达式）在构建 Excel 时转换为中文：

```java
sheetInfo.putFieldHandler("gender", (FieldValueConverter<Integer>) value -> 
                          value == null ? null : value.equals(1) ? "男生" : "女生");
```

完整的使用方法示例详见 [ExcelWriterTests.java](https://github.com/gaoice/easyexcel/blob/master/src/test/java/com/gaoice/easyexcel/test/writer/ExcelWriterTests.java) 。

#### 读 Excel

简单的读取为 `List<Map>` ：

```java
List<Map<Integer, Object>> result = ExcelReader.parseList("example.xlsx");
```

映射为 `List<Student>` ：

```java
BeanConfig<Student> beanConfig = new BeanConfig<Student>().setTargetClass(Student.class)
        .setFieldNames(new String[]{"name", "idcard", "gender", ...});

List<Student> result = ExcelReader.parseList("example.xlsx", beanConfig);
```

为 gender 字段设置转换器，把中文映射为实体类的 `Integer` 类型：

```java
beanConfig.putConverter("gender", wrapper ->
                wrapper.getStringValue() == null ? null : ("男生".equals(wrapper.getStringValue()) ? 1 : 0));
```

完整的使用方法示例详见 [ExcelReaderTests.java](https://github.com/gaoice/easyexcel/blob/master/src/test/java/com/gaoice/easyexcel/test/reader/ExcelReaderTests.java) 。