package com.file.util.support;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;


/**
 * @program: ExcelUtil
 * @description:
 * @author: zan.Kang
 * @create: 2019-06-04 09:03
 **/
public class SimpleSheets<T> {

    private Sheet sheet;


    public SimpleSheets(Sheet sheet) {
        this.sheet = sheet;
    }


    public List<T> convertWith(Integer rowIndex, Class<T> clazz) throws Exception {
        if (sheet == null) {
            return null;
        }
        return limit(rowIndex, sheet.getLastRowNum(), clazz);
    }


    public List<T> limitWith(Integer rowIndex, Integer rows, Class<T> clazz) throws Exception {
        return limit(rowIndex, rowIndex + rows, clazz);
    }

    private List<T> limit(Integer beginIndex, Integer endIndex, Class<T> clazz) throws Exception {
        List<T> resultRist = new ArrayList();
        for (int i = beginIndex; i < endIndex; i++) {
            T data = clazz.newInstance();
            Row r = sheet.getRow(i);
            if (r == null) {
                break;
            }
            Field[] fields = data.getClass().getDeclaredFields();
            for (int j = 0; j < fields.length; j++) {
                Cell cell = r.getCell(j);
                Field f = fields[j];
                f.setAccessible(true);
                if (f.getType() == Integer.class) {
                    f.set(data, Integer.parseInt(cell.getStringCellValue()));
                }
                System.out.println(f.getType().getName()+":equals:"+String.class.getName());
                if (f.getType().getName().equals(String.class.getName())) {
                    System.out.println("set");
                    f.set(data, cell.getStringCellValue());
                }
                resultRist.add(data);
            }
        }
        return resultRist;
    }


}
