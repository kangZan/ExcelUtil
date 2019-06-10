package com.file.util.support;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

import com.file.util.ExcelUtil;


/**
 * @program: ExcelUtil
 * @description:简单excel导出的poi支持
 * @author: zan.Kang
 * @create: 2019-05-22 16:25
 * @email: kangzans@163.com
 **/
public class SimpleExplorSupport implements ExcelUtil {

    private Collection[] datas;
    private String[] sheet_names;
    private ArrayList<String>[] header;
    private CellStyle[] datarow_styles;
    private CellStyle[] headerrow_style;
    private Integer sheet_num = 1;
    private Workbook workbook;
//    private Boolean acess_private;//是否只读私有属性

    /*公有方法**/

    /**
     * 构成excel的workbook并返回
     */
    public Workbook get() {
        createWorrkBooks();
        return workbook;
    }

    /**
     * 设置构成excel所需的属性
     */
    //设置多个sheet的表头行
    public void setHeaderRows(ArrayList<String>... header) {
        this.header = header;
    }

    //设置多个sheet的表头style
    //暂时使用默认表头样式@createDefaultDataCellStyle
//    public void setHeaderRowStyle(CellStyle... headerrow_style) {
//        this.headerrow_style = headerrow_style;
//    }

    //    //设置多个sheet的数据行
    public void setDataRows(Collection... datas) {
        this.datas = datas;
    }

    //设置多个sheet的数据行style
    //暂时使用默认数据行样式@createDefaultDataCellStyle
//    public void setDataRowStyle(CellStyle... datarow_styles) {
//        this.datarow_styles = datarow_styles;
//    }

    //设置多个sheet的sheet名字
    public void setSheetNames(String... sheet_names) {
        this.sheet_names = sheet_names;
    }

    //设置需要生成的sheet数
    public void setSheetNum(int num) {
        if (num > 0) sheet_num = num;
    }


    /*私有方法**/

    /**
     * 构建sheet
     */
    //根据属性信息生成workbook，并分别填入数据到data中
    private void createWorrkBooks() {
        workbook = new HSSFWorkbook();
        if (datas != null && datas.length > 0) {
            for (int i = 0; i < sheet_num; i++) {
                if (i + 1 > datas.length) {
                    break;
                }
                Sheet sheet = workbook.createSheet(getSheetName(i));
                createSheet(getHeaderRow(i), datas[i], sheet,
                        createDefaultHeadersCellStyle(workbook)//getHaderCellStyle(i)
                        , createDefaultDataCellStyle(workbook))// getDataCellStyle(i))
                ;

            }
        }
    }

    //创建单个sheet中的内容
    private void createSheet(Collection sheetHeaders, Collection sheetDatas, Sheet sheet, CellStyle headerCss, CellStyle dataCss) {
        Integer rowIndex = 0;
        rowIndex = pushDataToSheet(sheetHeaders, sheet, headerCss, rowIndex);//创建表头行
        setColWidth(sheet, sheetHeaders);
        pushDataToSheet(sheetDatas, sheet, dataCss, rowIndex);//创建数据行
    }

    private void setColWidth(Sheet sheet, Collection datas) {
        int i = 0;
        for (Object data : datas) {
            if (data != null) {
                sheet.setColumnWidth(i, data.toString().getBytes().length * 256);
            }
            i++;
        }
//                设置表头列宽
    }

    //放入data到sheet
    private Integer pushDataToSheet(Collection sheetDatas, Sheet sheet, CellStyle css, Integer rowIndex) {
        Collection<Object> cellDatas = null;
        for (Object rowData : sheetDatas) {
            if (rowData.getClass().getClassLoader() == null) {
                cellDatas = sheetDatas;//已经是最低粒度,直接作为填入单元格的数据（判断为表头）
                pushDataToRow(cellDatas, sheet.createRow(rowIndex), css);
                rowIndex++;
                break;
            } else {
                cellDatas = convertFieldToCellData(rowData); //解析对象属性，作为填入单元格的数据
                pushDataToRow(cellDatas, sheet.createRow(rowIndex), css);
            }
            rowIndex++;
        }
        return rowIndex;


    }


    //放入data到row
    private List<Object> convertFieldToCellData(Object rowData) {
        Class dataClass = rowData.getClass();
        Field[] fs = dataClass.getDeclaredFields();
        List<Object> cellDatas = new ArrayList<Object>();
        for (int i = 0; i < fs.length; i++) {
            /*设置只渲染非private字段，有点问题，待修改优化
             * getDeclaredFields   获取所有字段
             * getFields     获取公用字段
             * */
//            if (acess_private) {
//            fs[i].setAccessible(true)
//            } else {
//                if (!fs[i].isAccessible()) {
//                    System.out.println(fs[i].getName() + ":" + fs[i].isAccessible());
//                    continue;//跳过私有属性渲染
//                }
//            }
            fs[i].setAccessible(true); // 设置私有属性可以被访问的
            Object cellData = getFieldVal(fs[i], rowData);
            cellDatas.add(cellData);
        }
        return cellDatas;
    }


    private void pushDataToRow(Collection cellDatas, Row row, CellStyle css) {
        if (cellDatas == null || cellDatas.size() == 0) {
            return;
        }
        int i = 0;
        for (Object cellData : cellDatas) {
            pushValToCell(row.createCell(i), cellData, css);
            i++;
        }
    }


    //放入data.val到cell
    private void pushValToCell(Cell cell, Object cellData, CellStyle css) {
        if (cellData != null) {
            cell.setCellValue(cellData.toString());
        }
        if (css != null) {
            cell.setCellStyle(css);
        }
    }

    //获取属性值
    private Object getFieldVal(Field field, Object obj) {
        try {
            return field.get(obj);
        } catch (IllegalArgumentException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        return "此数据读取异常";
    }


    /**
     * 获取构成excel必要参数
     */

    //获取设置的sheet名字
    private String getSheetName(int index) {
        if (sheet_names != null && sheet_names.length > 0) {
            if (index + 1 <= sheet_names.length) {
                return sheet_names[index];
            }
        }
        return "Sheet" + index;
    }

    //获取设置的表头行样式
//    private CellStyle getHaderCellStyle(int index) {
//        if (headerrow_style != null) {
//            if (headerrow_style.length > 0) {
//                if (index + 1 <= headerrow_style.length) {
//                    return headerrow_style[index];
//                }
//            }
//            return headerrow_style[headerrow_style.length - 1];
//        }
//        return null;
//    }

    //获取设置的数据行样式
//    private CellStyle getDataCellStyle(int index) {
//        if (datarow_styles != null) {
//            if (datarow_styles.length > 0) {
//                if (index + 1 <= datarow_styles.length) {
//                    return datarow_styles[index];
//                }
//            }
//            return datarow_styles[datarow_styles.length - 1];
//        }
//        return null;
//    }

    //获取设置的表头数据
    private ArrayList getHeaderRow(int index) {
        if (header != null) {
            if (header.length > 0) {
                if (index + 1 <= header.length) {
                    return header[index];
                }
            }
            return header[header.length - 1];
        }
        return null;
    }

    //创建数据单元格的样式
    private static CellStyle createDefaultDataCellStyle(Workbook workbook) {
        CellStyle css = workbook.createCellStyle();
        css.setBorderBottom(BorderStyle.THIN);
        css.setBorderTop(BorderStyle.THIN);
        css.setBorderLeft(BorderStyle.THIN);
        css.setBorderRight(BorderStyle.THIN);
        css.setAlignment(HorizontalAlignment.CENTER); // 居中
        css.setVerticalAlignment(VerticalAlignment.CENTER); // 居中
        return css;
    }

    //创建表头单元格的样式
    private static CellStyle createDefaultHeadersCellStyle(Workbook workbook) {
        CellStyle headercss = createDefaultDataCellStyle(workbook);
        headercss.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        headercss.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return headercss;
    }
}
