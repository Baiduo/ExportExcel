package com.example.baiduo.exceltest;

import android.Manifest;
import android.annotation.SuppressLint;
import android.content.Intent;
import android.os.Environment;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.text.TextUtils;
import android.view.View;
import android.widget.EditText;
import android.widget.TextView;

import com.tbruyelle.rxpermissions2.RxPermissions;


import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;

import butterknife.BindView;
import butterknife.ButterKnife;
import io.reactivex.functions.Consumer;
import jxl.format.Alignment;
import jxl.format.Colour;
import jxl.format.VerticalAlignment;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WriteException;

public class MainActivity extends AppCompatActivity {

    List<Object> list = new ArrayList();

    private static final String DEFAULT_TABLE_NAME = "工作表";
    private static final String DEFAULT_NAME = "孙悟空";
    private static final String DEFAULT_AGE = "502";
    private static final String DEFAULT_WORK = "偷桃";
    private static final int DEFAULT_ITEM_COUNT = 10;


    @BindView(R.id.et_table_name)
    EditText etTableName;
    @BindView(R.id.et_count)
    EditText etCount;
    @BindView(R.id.et_name)
    EditText etName;
    @BindView(R.id.et_age)
    EditText etAge;
    @BindView(R.id.et_work)
    EditText etWork;
    @BindView(R.id.tv_path)
    TextView tv_path;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        ButterKnife.bind(this);


    }

    @SuppressLint("CheckResult")
    public void addTable(View view) {
        final String tableName = etTableName.getText().toString().trim();
        final String count = etCount.getText().toString();
        final String name = etName.getText().toString();
        final String age = etAge.getText().toString();
        final String work = etWork.getText().toString();
        RxPermissions rxPermissions = new RxPermissions(this);
        rxPermissions.request(Manifest.permission.WRITE_EXTERNAL_STORAGE)
                .subscribe(new Consumer<Boolean>() {
                    @Override
                    public void accept(Boolean aBoolean) {
                        if (aBoolean) {
                            exportExcel(tableName,count,name,age,work);
                        }
                    }
                });
    }

    /**
     * 导出数据,耗时操作最好放在子线程,因为是demo暂时不做此操作
     * @param tableName
     * @param count
     * @param name
     * @param age
     * @param work
     */
    private void exportExcel(String tableName, String count, String name, String age, String work) {
        ExcelUtils excelUtils = ExcelUtils.getInstance().create(Environment.getExternalStorageDirectory().getAbsolutePath() + "/ExportExcel", TextUtils.isEmpty(tableName)?DEFAULT_TABLE_NAME:tableName);
        int size = TextUtils.isEmpty(count)?DEFAULT_ITEM_COUNT:Integer.parseInt(count);
        String itemName = TextUtils.isEmpty(name) ? DEFAULT_NAME : name;
        String itemAge = TextUtils.isEmpty(age) ? DEFAULT_AGE : age;
        String itemWork = TextUtils.isEmpty(work) ? DEFAULT_WORK : work;
        for (int x = 0; x < size; x++) {
            JavaBean javaBean = new JavaBean(itemName, itemAge, itemWork);
            list.add(javaBean);
        }
        try {
            WritableCellFormat titleFormat = new WritableCellFormat();
            WritableCellFormat dataFormat = new WritableCellFormat();
            titleFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
            titleFormat.setAlignment(Alignment.CENTRE);
            dataFormat.setAlignment(Alignment.RIGHT);
            excelUtils.createSheetSetTitle(tableName, new String[]{"姓名", "年龄", "工作"}, titleFormat)
                    .fillData(list, dataFormat).close();
            tv_path.setText("表格已存入: "+Environment.getExternalStorageDirectory().getAbsolutePath() + "/ExportExcel");

        } catch (WriteException e) {
            e.printStackTrace();
        }
    }

}
