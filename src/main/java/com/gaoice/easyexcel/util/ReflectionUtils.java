package com.gaoice.easyexcel.util;

import org.apache.commons.lang3.StringUtils;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Modifier;
import java.util.List;

/**
 * 反射工具类
 *
 * @author gaoice
 */
public class ReflectionUtils {

    public static Field[] getFieldChain(Class<?> clazz, List<String> fieldNameChain) throws NoSuchFieldException {
        Assert.notEmpty(fieldNameChain, "fieldNameChain must be non-empty");
        return getFieldChain(clazz, fieldNameChain.toArray(new String[0]));
    }

    public static Field[] getFieldChain(Class<?> clazz, String fieldNameChainStr) throws NoSuchFieldException {
        Assert.notBlank(fieldNameChainStr, "fieldNameChainStr must be not blank");
        return getFieldChain(clazz, StringUtils.split(fieldNameChainStr, "."));
    }

    public static Object getFieldValue(Object object, String fieldNameChainStr) throws NoSuchFieldException, IllegalAccessException {
        Assert.notNull(object, "object must be non-null");
        return getFieldValue(object, getFieldChain(object.getClass(), fieldNameChainStr));
    }

    public static Object getFieldValue(Object object, String[] fieldNameChain) throws NoSuchFieldException, IllegalAccessException {
        Assert.notNull(object, "object must be non-null");
        return getFieldValue(object, getFieldChain(object.getClass(), fieldNameChain));
    }

    public static Object getFieldValue(Object object, List<Field> fieldChain) throws IllegalAccessException {
        Assert.notEmpty(fieldChain, "fieldChain must be non-empty");
        return getFieldValue(object, fieldChain.toArray(new Field[0]));
    }

    public static void setFieldValue(Object object, String fieldNameChainStr, Object value) throws NoSuchFieldException, InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException {
        Assert.notNull(object, "object must be non-null");
        setFieldValue(object, getFieldChain(object.getClass(), fieldNameChainStr), value);
    }

    public static void setFieldValue(Object object, String[] fieldNameChain, Object value) throws NoSuchFieldException, InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException {
        Assert.notNull(object, "object must be non-null");
        setFieldValue(object, getFieldChain(object.getClass(), fieldNameChain), value);
    }

    public static void setFieldValue(Object object, List<Field> fieldChain, Object value) throws NoSuchFieldException, InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException {
        Assert.notEmpty(fieldChain, "fieldChain must be non-empty");
        setFieldValue(object, fieldChain.toArray(new Field[0]), value);
    }

    /**
     * 获取连续的字段数组
     *
     * @param clazz          class
     * @param fieldNameChain 连续的字段名数组
     * @return 连续的 Field 数组
     */
    public static Field[] getFieldChain(Class<?> clazz, String[] fieldNameChain) throws NoSuchFieldException {
        Assert.notNull(clazz, "clazz must be non-null");
        Assert.notEmpty(fieldNameChain, "fieldNameChain must be non-empty");
        Field[] fieldChain = new Field[fieldNameChain.length];
        int i = 0;
        for (; i < fieldNameChain.length; i++) {
            Field field = clazz.getDeclaredField(fieldNameChain[i]);
            makeAccessible(field);
            fieldChain[i] = field;
            clazz = field.getType();
        }
        return fieldChain;
    }

    public static Object getFieldValue(Object object, Field[] fieldChain) throws IllegalAccessException {
        Assert.notNull(object, "object must be non-null");
        Assert.notEmpty(fieldChain, "fieldChain must be non-empty");
        Object o = object;
        for (Field f : fieldChain) {
            makeAccessible(f);
            if (o != null) {
                o = f.get(o);
            } else {
                break;
            }
        }
        return o;
    }

    public static void setFieldValue(Object object, Field[] fieldChain, Object value) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException, InstantiationException {
        Assert.notNull(object, "object must be non-null");
        Assert.notEmpty(fieldChain, "fieldChain must be non-empty");
        int i = 0;
        for (; i < fieldChain.length - 1; i++) {
            Field f = fieldChain[i];
            Object o = f.get(object);
            if (o == null) {
                o = f.getType().getDeclaredConstructor().newInstance();
                f.set(object, o);
            }
            object = o;
        }
        Field targetField = fieldChain[i];
        makeAccessible(targetField);
        targetField.set(object, value);
    }

    @SuppressWarnings("deprecation")  // on JDK 9
    public static void makeAccessible(Field field) {
        if ((!Modifier.isPublic(field.getModifiers()) ||
                !Modifier.isPublic(field.getDeclaringClass().getModifiers()) ||
                Modifier.isFinal(field.getModifiers())) && !field.isAccessible()) {
            field.setAccessible(true);
        }
    }
}
