package com.gaoice.easyexcel.test.entity;

/**
 * 性别统计结果
 *
 * @author gaoice
 */
public class GenderCountResult {

    int manNum = 0;
    int womanNum = 0;

    @Override
    public String toString() {
        return "男生" + manNum + "人，女生" + womanNum + "人";
    }

    public void addManNum() {
        manNum++;
    }

    public void addWomanNum() {
        womanNum++;
    }

}
