package com.gaoice.easyexcel.test;

import com.gaoice.easyexcel.ExcelBuilder;
import com.gaoice.easyexcel.data.Converter;
import org.junit.Before;
import org.junit.Test;

import java.io.FileOutputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @since 1.1 支持对 List<Map<?,?>> 类型的数据进行处理
 */
public class MapListTests {

    private List<Map<String, Object>> mapList;

    @Before
    public void mock() {
        Map<String, Object> user1 = new HashMap<>();
        user1.put("name", "用户1");
        user1.put("level", 5);
        user1.put("lastLogin", LocalDateTime.now());

        Map<String, Object> user2 = new HashMap<>();
        user2.put("name", "用户2");
        user2.put("level", 3);
        user2.put("lastLogin", null);

        mapList = new ArrayList<>(2);
        mapList.add(user1);
        mapList.add(user2);
    }

    /**
     * 支持 List<Map<?, ?>>
     * <p>
     * 单个 sheet 支持 链式调用
     */
    @Test
    public void mapList() throws Exception {
        String sheetName = "用户表";
        String title = "用户信息";
        String[] columnNames = {"用户名", "等级", "上次登录时间"};
        String[] fieldNames = {"name", "level", "lastLogin"};

        FileOutputStream file = new FileOutputStream("mapList.xlsx");

        ExcelBuilder.builder()
                .setSheetName(sheetName)
                .setTitle(title)
                .setColumnNames(columnNames)
                .setFieldNames(fieldNames)
                .setList(mapList)
                .putConverter(fieldNames[2], (Converter<LocalDateTime, String>) (sheetInfo1, value, listIndex, columnIndex) -> {
                    if (value != null) {
                        return DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss").format(value);
                    }
                    return "未知";
                })
                .build(file);

        file.close();
    }
}
