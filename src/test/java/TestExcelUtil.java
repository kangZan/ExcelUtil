import java.io.FileOutputStream;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Workbook;

import com.file.util.support.SimplePOISupport;

/**
 * @program: ExcelUtil
 * @description:
 * @author: zan.Kang
 * @create: 2019-05-23 16:49
 **/
public class TestExcelUtil {

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
        Student student = new Student();
        student.setName("Holden");
        student.setAge(24);
        student.setSchool("云梦中学");
        laa[1] = new ArrayList();
        Student student2 = new Student();
        student2.setName("子孑");
        student2.setAge(24);
        student2.setSchool("湖南理工");
        laa[0].add(student2);
        laa[0].add(student);
        laa[1].add(student);
        laa[1].add(student2);
        return laa;
    }

    public static void main(String[] args) throws Exception {
        ArrayList[] headers = initHeaderArray();
        String[] sheetnames = {"保信1"};
        ArrayList[] datas = initDataArray();
        SimplePOISupport poiUtil = new SimplePOISupport();
        poiUtil.setHeaderRows(headers);//设置多个sheet表头行
        poiUtil.setDataRows(datas);//设置多个sheet数据行
        poiUtil.setSheetNames(sheetnames);//设置sheet名
        poiUtil.setSheetNum(2);//设置sheet数

//        poiUtil.setDataRowStyle(null);//设置表头样式
//        poiUtil.setHeaderRowStyle(null);//设置表头样式
        Workbook wb = poiUtil.get();

        FileOutputStream f = new FileOutputStream("test.xlsx");
        wb.write(f);
        f.close();
        wb.close();
    }


}

class Student {

    private String name;
    private Integer age;
    private String school;

    public void setName(String name) {
        this.name = name;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public void setSchool(String school) {
        this.school = school;
    }

    public String getName() {
        return name;
    }

    public Integer getAge() {
        return age;
    }

    public String getSchool() {
        return school;
    }
}
