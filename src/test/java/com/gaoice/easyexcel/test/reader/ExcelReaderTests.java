package com.gaoice.easyexcel.test.reader;

import com.gaoice.easyexcel.reader.ExcelReader;
import com.gaoice.easyexcel.reader.SheetResult;
import com.gaoice.easyexcel.reader.converter.CellConverter;
import com.gaoice.easyexcel.reader.sheet.BeanConfig;
import com.gaoice.easyexcel.test.entity.Student;
import org.junit.Test;

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Map;

/**
 * @author gaoice
 */
public class ExcelReaderTests {

    /**
     * 读取为 map
     */
    @Test
    public void readMap() throws IOException {
        SheetResult<Map<Integer, Object>> result = ExcelReader.parseSheetResult("1-simple.xlsx");

        assert result.getResult().size() == 10000;
    }

    /**
     * 简单的读取
     * 设置忽略字段映射的异常，所以 "gender", "birthday" 会读取失败但并不会抛异常
     * 基本类型的包装类在 CellConverterRegistry 中设置的有默认的转换器
     */
    @Test
    public void simpleRead() throws Exception {
        BeanConfig<Student> beanConfig = new BeanConfig<Student>().setTargetClass(Student.class)
                //设置 字段
                .setFieldNames(new String[]{"name", "cardId", "gender", "birthday", "grade.chineseGrade", "grade.mathGrade", "grade.englishGrade"})
                //忽略字段映射的异常
                .setIgnoreException(true);

        SheetResult<Student> result = ExcelReader.parseSheetResult("1-simple.xlsx", beanConfig);

        assert result.getResult().size() == 10000;
    }

    @Test
    public void read() throws Exception {
        //使用空字符串忽略第一列序号列
        String[] fieldNames = {"name", "cardId", "gender", "birthday", "grade.chineseGrade", "grade.mathGrade", "grade.englishGrade"};

        SheetResult<Student> result = ExcelReader.parseSheetResult("3-complexHandler.xlsx",
                new BeanConfig<Student>().setTargetClass(Student.class)
                        .setFieldNames(fieldNames)
                        .setIgnoreException(false)
                        //倒数第一行是统计结果，设置不读取倒数第一行
                        .setListLastRowNum(-1)
                        //设置性别列的转换器
                        .putConverter("gender", GENDER_CONVERTER)
                        //设置日期转换器
                        .putConverter("birthday", DATE_CONVERTER)
        );

        assert result.getResult().size() == 10000;
    }

    public static CellConverter DATE_CONVERTER = context -> {
        try {
            String value = context.getStringValue();
            return value == null ? null : new SimpleDateFormat("yyyy-MM-dd").parse(value);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    };

    public static CellConverter GENDER_CONVERTER = context -> {
        String value = context.getStringValue();
        if ("男生".equals(value)) {
            return 1;
        }
        if ("女生".equals(value)) {
            return 0;
        }
        return null;
    };
}
