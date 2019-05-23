package com.file.util;

import java.util.ArrayList;
import java.util.Collection;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @program: bx-station
 * @description:
 * @author: zan.Kang
 * @create: 2019-05-22 16:25
 * @email: kangzans@163.com
 **/
public interface ExcelUtil {

    /*
     *获取最终输出workbook
     **/
    Workbook get();
}
