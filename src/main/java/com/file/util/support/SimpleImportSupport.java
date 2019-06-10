package com.file.util.support;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @program: ExcelUtil
 * @description:简单excel导出的hutool支持
 * @author: zan.Kang
 * @create: 2019-05-24 08:38
 **/
public class SimpleImportSupport {

    private List<SimpleSheets> simpleSheets;

    public SimpleImportSupport setWorkbook(Workbook workbook) {
        simpleSheets = new ArrayList<SimpleSheets>();
        Iterator<Sheet> iterator = workbook.sheetIterator();
        if (iterator != null) {
            while (iterator.hasNext()) {
                simpleSheets.add(new SimpleSheets(iterator.next()));
            }
            return this;
        }
        return null;
    }


    public SimpleSheets get(int i) {
        return simpleSheets.get(i);
    }


}



