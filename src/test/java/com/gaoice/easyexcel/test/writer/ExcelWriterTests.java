package com.gaoice.easyexcel.test.writer;

import com.gaoice.easyexcel.test.entity.GenderCountResult;
import com.gaoice.easyexcel.test.entity.Grade;
import com.gaoice.easyexcel.test.entity.Student;
import com.gaoice.easyexcel.test.writer.style.MySheetStyle;
import com.gaoice.easyexcel.writer.ExcelWriter;
import com.gaoice.easyexcel.writer.SheetInfo;
import com.gaoice.easyexcel.writer.handler.FieldHandler;
import com.gaoice.easyexcel.writer.handler.FieldHandlers;
import com.gaoice.easyexcel.writer.handler.FieldValueConverter;
import org.junit.Before;
import org.junit.Test;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ExcelWriterTests {

    List<Student> studentList;

    /**
     * 生成测试数据 studentList
     */
    @Before
    public void initList() {
        studentList = new ArrayList<>();
        for (int i = 0; i < 10000; i++) {
            studentList.add(genStudent(i));
        }
    }

    /**
     * 最简单的使用
     */
    @Test
    public void simple() throws Exception {
        // 1.定义 excel 每一列对应的实体类字段
        String[] classFieldNames = {"name", "cardId", "gender", "birthday", "grade.chineseGrade", "grade.mathGrade", "grade.englishGrade"};
        // 2.直接写入文件并关闭输出流
        ExcelWriter.writeAndCloseOutputStream(new SheetInfo(classFieldNames, studentList), new FileOutputStream("1-simple.xlsx"));
    }

    /**
     * 在 simple 的基础上设置 sheetName，标题，列名
     * 为 性别字段 这一列设置处理器，把数字 0，1 转换为汉字
     */
    @Test
    public void simpleHandler() throws Exception {
        String sheetName = "学生信息表";
        String title = "学生信息表";
        String[] columnNames = {"姓名", "身份证号", "性别", "出生日期", "语文分数", "数学分数", "英语分数"};
        String[] classFieldNames = {"name", "cardId", "gender", "birthday", "grade.chineseGrade", "grade.mathGrade", "grade.englishGrade"};
        SheetInfo sheetInfo = new SheetInfo(sheetName, title, columnNames, classFieldNames, studentList);
        /*
         * 为 gender 字段设置处理器，将 性别 0，1 转换为 女生，男生
         * FieldValueConverter 是 FieldHandler 简单的包装，为了使值转换功能的 lambda 表达式更简单直观
         */
        sheetInfo.putFieldHandler("gender",
                (FieldValueConverter<Integer>) value -> value == null ? null : value.equals(1) ? "男生" : "女生");

        ExcelWriter.writeAndCloseOutputStream(sheetInfo, new FileOutputStream("2-simpleHandler.xlsx"));
    }

    /**
     * 具有 转换和统计 功能稍复杂的 handler 示例
     * 性别列增加统计功能对 男女人数 进行统计
     */
    @Test
    public void complexHandler() throws Exception {
        String sheetName = "计算机学院";
        String title = "学生信息表";
        String[] columnNames = {"姓名", "身份证号", "性别", "出生日期", "语文分数", "数学分数", "英语分数"};
        String[] classFieldNames = {"name", "cardId", "gender", "birthday", "grade.chineseGrade", "grade.mathGrade", "grade.englishGrade"};
        SheetInfo sheetInfo = new SheetInfo(sheetName, title, columnNames, classFieldNames, studentList);
        /*
         * GENDER_HANDLER 实现类性别的转换和统计功能
         * 只要在 FieldHandler 中调用了 context.setConvertedValue(Object) 方法设置非 null 的统计结果，生成的 sheet 就会追加统计结果行
         */
        sheetInfo.putFieldHandler("gender", GENDER_HANDLER);
        /*
         * 为 birthday 字段设置处理器，覆盖 FieldHandlerRegistry 中 Date 类型的默认的转换器，达到格式化日期为 yyyy-MM-dd 格式的目的
         * 如果要替换全局所有 Date 类型默认的格式，使用 FieldHandlerRegistry.register(Class, FieldHandler) 方法覆盖默认的处理器即可
         */
        sheetInfo.putFieldHandler("birthday",
                (FieldValueConverter<Date>) value -> value == null ? null : new SimpleDateFormat("yyyy-MM-dd").format(value));

        ExcelWriter.writeAndCloseOutputStream(sheetInfo, new FileOutputStream("3-complexHandler.xlsx"));
    }

    /**
     * 在 complexHandler 的基础上
     * 使用 # 开头的虚拟字段定义 序号列 总分数列，并为虚拟字段添加处理器
     */
    @Test
    public void virtualField() throws Exception {
        String sheetName = "学生信息表";
        String title = "学生信息表";
        String[] columnNames = {"序号", "姓名", "身份证号", "性别", "出生日期", "语文分数", "数学分数", "英语分数", "总分数"};
        String[] classFieldNames = {"#order", "name", "cardId", "gender", "birthday", "grade.chineseGrade", "grade.mathGrade", "grade.englishGrade", "#totalGrade"};
        SheetInfo sheetInfo = new SheetInfo(sheetName, title, columnNames, classFieldNames, studentList);

        sheetInfo.putFieldHandler("gender", GENDER_CONVERTER)
                // 虚拟字段 #order，通过转换器生成序号，虚拟字段的泛型都是 SheetInfo.list 的元素类型
                .putFieldHandler("#order", FieldHandlers.ORDER_GENERATOR)
                // 虚拟字段 #totalGrade，通过转换器生成总分数
                .putFieldHandler("#totalGrade", TOTAL_GRADE_CONVERTER);

        ExcelWriter.writeAndCloseOutputStream(sheetInfo, new FileOutputStream("4-virtualField.xlsx"));
    }

    /**
     * 设置自定义style，对60分以下的分数标红，纯数字列自适应会遮挡，对身份证长度进行调整
     */
    @Test
    public void sheetStyle() throws Exception {
        String sheetName = "学生信息表";
        String title = "学生信息表";
        String[] columnNames = {"序号", "姓名", "身份证号", "性别", "出生日期", "语文分数", "数学分数", "英语分数", "总分数"};
        String[] classFieldNames = {"#order", "name", "cardId", "gender", "birthday", "grade.chineseGrade", "grade.mathGrade", "grade.englishGrade", "#totalGrade"};
        SheetInfo sheetInfo = new SheetInfo(sheetName, title, columnNames, classFieldNames, studentList);
        sheetInfo.putFieldHandler("gender", GENDER_HANDLER)
                .putFieldHandler("#order", FieldHandlers.ORDER_GENERATOR_RESULT_TEXT)
                .putFieldHandler("#totalGrade", TOTAL_GRADE_CONVERTER);

        // 自定义样式，对分数列小于60分的分数设置为红色字体
        sheetInfo.setSheetStyle(new MySheetStyle());

        ExcelWriter.writeAndCloseOutputStream(sheetInfo, new FileOutputStream("5-sheetStyle.xlsx"));
    }

    /**
     * 两个sheet
     */
    @Test
    public void twoSheet() throws Exception {
        String sheetName1 = "学生信息表";
        String title1 = "学生信息表";
        String[] columnNames1 = {"序号", "姓名", "身份证号", "性别", "出生日期"};
        String[] classFieldNames1 = {"#order", "name", "cardId", "gender", "birthday"};
        SheetInfo sheetInfo1 = new SheetInfo(sheetName1, title1, columnNames1, classFieldNames1, studentList);
        sheetInfo1.putFieldHandler("#order", FieldHandlers.ORDER_GENERATOR)
                .putFieldHandler("gender", GENDER_CONVERTER)
                .setSheetStyle(new MySheetStyle());

        String sheetName2 = "学生分数表";
        String title2 = "学生分数表";
        String[] columnNames2 = {"序号", "姓名", "语文分数", "数学分数", "英语分数", "总分数"};
        String[] classFieldNames2 = {"#order", "name", "grade.chineseGrade", "grade.mathGrade", "grade.englishGrade", "#totalGrade"};
        SheetInfo sheetInfo2 = new SheetInfo(sheetName2, title2, columnNames2, classFieldNames2, studentList);
        sheetInfo2.putFieldHandler("#order", FieldHandlers.ORDER_GENERATOR)
                .putFieldHandler("#totalGrade", TOTAL_GRADE_CONVERTER)
                .setSheetStyle(new MySheetStyle());

        List<SheetInfo> sheetInfos = new ArrayList<>();
        sheetInfos.add(sheetInfo1);
        sheetInfos.add(sheetInfo2);

        ExcelWriter.writeAndCloseOutputStream(sheetInfos, new FileOutputStream("6-twoSheet.xlsx"));
    }

    /**
     * FieldValueConverter 接口使用示例，继承 FieldHandler ，只关注字段值的转换，lambda 表达式更简洁
     */
    static FieldValueConverter<Integer> GENDER_CONVERTER = value ->
            value == null ? null : value.equals(1) ? "男生" : "女生";

    /**
     * FieldHandler 接口使用示例，可进行转换 统计 设置单元格等操作
     * 此处实现 性别 转换 + 统计
     * GenderCountResult 保存了统计结果，使用自定义结果类型要注意重写 toString 方法，以便在统计行显示自己想要的结果
     */
    static FieldHandler<Integer> GENDER_HANDLER = (context) -> {
        GenderCountResult result = (GenderCountResult) context.getCountedResult();
        if (result == null) {
            result = new GenderCountResult();
        }
        Integer v = context.getValue();
        String s = null;
        if (v != null) {
            if (v.equals(1)) {
                result.addManNum();
                s = "男生";
            } else {
                result.addWomanNum();
                s = "女生";
            }
        }
        context.setConvertedValue(s);
        context.setCountedResult(result);
    };

    /**
     * 总分数转换器，对语数外三门分数进行求和
     * 总分数是虚拟字段
     * 虚拟字段的类型是 {@link SheetInfo#getList()} 元素的类型
     */
    static FieldValueConverter<Student> TOTAL_GRADE_CONVERTER = (student) -> {
        if (student == null) {
            return null;
        }
        Grade grade = student.getGrade();
        if (grade == null) {
            return null;
        }
        Double result = (double) 0;
        if (grade.getChineseGrade() != null) {
            result += grade.getChineseGrade();
        }
        if (grade.getMathGrade() != null) {
            result += grade.getMathGrade();
        }
        if (grade.getEnglishGrade() != null) {
            result += grade.getEnglishGrade();
        }
        return result;
    };


    String[] chinese = {"零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖", "拾"};

    private Student genStudent(int i) {
        Student student = new Student();
        student.setName("张" + chinese[i % 11]);
        student.setCardId("123123123412341234");
        student.setGender(i % 2);
        student.setBirthday(new Date());
        Grade grade = new Grade();
        grade.setChineseGrade(i % 10 * 10 + i % 10);
        grade.setMathGrade((float) (i % 10 * 10 + i % 10));
        grade.setEnglishGrade((double) (i % 10 * 10 + i % 10));
        student.setGrade(grade);
        return student;
    }
}
