package com.example.baiduo.exceltest;

import java.io.Serializable;

/**
 * 反射后会根据字母顺序进行排序
 */
public class JavaBean implements Serializable {
    private String c_work;
    private String a_name;
    private String b_age;

    JavaBean(String name,String age,String work){
        this.a_name = name;
        this.b_age = age;
        this.c_work = work;
    }

    public String getC_work() {
        return c_work;
    }

    public void setC_work(String c_work) {
        this.c_work = c_work;
    }

    public String getA_name() {
        return a_name;
    }

    public void setA_name(String a_name) {
        this.a_name = a_name;
    }

    public String getB_age() {
        return b_age;
    }

    public void setB_age(String b_age) {
        this.b_age = b_age;
    }
}
