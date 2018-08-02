package com.example.baiduo.exceltest

import java.io.Serializable

/**
 * 反射后会根据字母顺序进行排序
 */
class JavaBean internal constructor(var a_name: String?, var b_age: String?, var c_work: String?) : Serializable
