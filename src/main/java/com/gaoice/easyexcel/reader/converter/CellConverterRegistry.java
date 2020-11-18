package com.gaoice.easyexcel.reader.converter;

import java.util.Optional;
import java.util.concurrent.ConcurrentHashMap;

/**
 * @author gaoice
 */
public class CellConverterRegistry {
    private static final ConcurrentHashMap<Class<?>, CellConverter> CONVERTER_MAP = new ConcurrentHashMap<>();

    static {
        register(Integer.class, CellConverters.INTEGER);
        register(int.class, CellConverters.INTEGER);
        register(Double.class, CellConverters.DOUBLE);
        register(double.class, CellConverters.DOUBLE);
        register(Float.class, CellConverters.FLOAT);
        register(float.class, CellConverters.FLOAT);
        register(Long.class, CellConverters.LONG);
        register(long.class, CellConverters.LONG);
        register(Short.class, CellConverters.SHORT);
        register(short.class, CellConverters.SHORT);
        register(Boolean.class, CellConverters.BOOLEAN);
        register(boolean.class, CellConverters.BOOLEAN);
        register(Character.class, CellConverters.CHAR);
        register(char.class, CellConverters.CHAR);
        register(Byte.class, CellConverters.BYTE);
        register(byte.class, CellConverters.BYTE);
    }

    public static void register(Class<?> clazz, CellConverter converter) {
        CONVERTER_MAP.put(clazz, converter);
    }

    public static void remove(Class<?> clazz) {
        CONVERTER_MAP.remove(clazz);
    }

    public static Optional<CellConverter> get(Class<?> clazz) {
        return Optional.ofNullable(CONVERTER_MAP.get(clazz));
    }

    public static void clear() {
        CONVERTER_MAP.clear();
    }
}
