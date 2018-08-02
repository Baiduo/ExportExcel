package com.example.baiduo.exceltest

import android.Manifest
import android.annotation.SuppressLint
import android.content.Intent
import android.os.Environment
import android.support.v7.app.AppCompatActivity
import android.os.Bundle
import android.text.TextUtils
import android.view.View
import android.widget.EditText
import android.widget.TextView

import com.tbruyelle.rxpermissions2.RxPermissions


import java.io.IOException
import java.util.ArrayList
import java.util.LinkedHashMap

import butterknife.BindView
import io.reactivex.functions.Consumer
import jxl.format.Alignment
import jxl.format.Colour
import jxl.format.VerticalAlignment
import jxl.write.WritableCellFormat
import jxl.write.WritableFont
import jxl.write.WriteException
import kotlinx.android.synthetic.main.activity_main.*

class MainActivity : AppCompatActivity() {

    private var list: MutableList<Any> = ArrayList()


    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)


    }

    fun addTable(view: View?) {
        val tableName = et_table_name.text.toString()
        val count = et_count.text.toString()
        val name = et_name.text.toString()
        val age = et_age.text.toString()
        val work = et_work.text.toString()
        val rxPermissions = RxPermissions(this)
        rxPermissions.request(Manifest.permission.WRITE_EXTERNAL_STORAGE)
                .subscribe { aBoolean ->
                    if (aBoolean!!) {
                        exportExcel(tableName, count, name, age, work)
                    }
                }
    }

    /**
     * 导出数据,耗时操作最好放在子线程,因为是demo暂时不做此操作
     * @param tableName
     * @param count
     * @param name
     * @param age
     * @param work
     */
    private fun exportExcel(tableName: String, count: String, name: String, age: String, work: String) {
        val excelUtils = ExcelUtils.getInstance().create(Environment.getExternalStorageDirectory().absolutePath + "/ExportExcel", if (TextUtils.isEmpty(tableName)) DEFAULT_TABLE_NAME else tableName)
        val size = if (TextUtils.isEmpty(count)) DEFAULT_ITEM_COUNT else Integer.parseInt(count)
        val itemName = if (TextUtils.isEmpty(name)) DEFAULT_NAME else name
        val itemAge = if (TextUtils.isEmpty(age)) DEFAULT_AGE else age
        val itemWork = if (TextUtils.isEmpty(work)) DEFAULT_WORK else work
        for (x in 0 .. size) {
            val javaBean = JavaBean(itemName, itemAge, itemWork)
            list.add(javaBean)
        }
        try {
            val titleFormat = WritableCellFormat()
            val dataFormat = WritableCellFormat()
            titleFormat.verticalAlignment = VerticalAlignment.CENTRE
            titleFormat.alignment = Alignment.CENTRE
            dataFormat.alignment = Alignment.RIGHT
            excelUtils.createSheetSetTitle(tableName, arrayOf("姓名", "年龄", "工作"), titleFormat)
                    .fillData(list, dataFormat).close()
            tv_path!!.text = "表格已存入: " + Environment.getExternalStorageDirectory().absolutePath + "/ExportExcel"

        } catch (e: WriteException) {
            e.printStackTrace()
        }

    }

    companion object {

        private val DEFAULT_TABLE_NAME = "工作表"
        private val DEFAULT_NAME = "孙悟空"
        private val DEFAULT_AGE = "502"
        private val DEFAULT_WORK = "偷桃"
        private val DEFAULT_ITEM_COUNT = 10
    }

}
