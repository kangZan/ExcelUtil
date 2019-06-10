import java.io.FileOutputStream;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Workbook;

import com.file.util.support.SimpleExplorSupport;

/**
 * @program: ExcelUtil
 * @description:
 * @author: zan.Kang
 * @create: 2019-05-23 16:49
 **/
public class TestExcelUtilExplor {
    public static ArrayList[] initHeaderArray() {
        ArrayList[] laa = new ArrayList[2];
        laa[0] = new ArrayList();
        laa[0].add("姓名");
        laa[0].add("年龄");
        laa[0].add("学校");
        laa[1] = new ArrayList();
        laa[1].add("姓名1");
        laa[1].add("年龄1");
        laa[1].add("学校1");
        return laa;
    }
    public static ArrayList[] initDataArray() {
        ArrayList[] laa = new ArrayList[2];
        laa[0] = new ArrayList();
        model.Student student = new model.Student();
        student.setName("Holden");
        student.setAge(24);
        student.setSchool("云梦中学");
        laa[1] = new ArrayList();
        model.Student student2 = new model.Student();
        student2.setName("子孑");
        student2.setAge(22);
        student2.setSchool("湖南理工");
        laa[0].add(student2);
        laa[0].add(student);
        laa[1].add(student);
        laa[1].add(student2);
        return laa;
    }
    public static void main(String[] args) throws Exception {
        /*待导出的数据*/
        ArrayList[] headers = initHeaderArray();
        String[] sheetnames = {"Sheet123"};
        ArrayList[] datas = initDataArray();
        /*set数据到工具类*/
        SimpleExplorSupport poiUtil = new SimpleExplorSupport();
        poiUtil.setHeaderRows(headers);//设置多个sheet表头行
        poiUtil.setDataRows(datas);//设置多个sheet数据行
        poiUtil.setSheetNames(sheetnames);//设置sheet名
        poiUtil.setSheetNum(2);//设置sheet数
        /*获取workbook并导出*/
        Workbook wb = poiUtil.get();
        FileOutputStream f = new FileOutputStream("test.xlsx");
        wb.write(f);
        f.close();
        wb.close();
    }
}