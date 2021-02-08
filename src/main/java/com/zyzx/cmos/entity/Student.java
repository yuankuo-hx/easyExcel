package com.zyzx.cmos.entity;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;
import lombok.Data;
import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;

import java.io.Serializable;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class Student extends BaseRowModel implements Serializable {

    @ExcelIgnore
    private String vid;

    @ExcelProperty(value = "学号",index = 0)
    private String studentId;

    @ExcelProperty(value = "姓名",index = 1)
    private String name;

    @ExcelProperty(value = "性别",index = 2)
    private String sex;


}
