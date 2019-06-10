import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import com.file.util.support.SimpleExplorSupport;
import com.file.util.support.SimpleImportSupport;
import com.sun.corba.se.spi.orbutil.threadpool.Work;

import model.Student;

/**
 * @program: ExcelUtil
 * @description:
 * @author: zan.Kang
 * @create: 2019-05-23 16:49
 **/
public class TestExcelUtilImport {

    public static void main(String[] args) throws Exception {
        SimpleImportSupport poiUtil = new SimpleImportSupport();

        FileInputStream f = new FileInputStream("test.xlsx");
        Workbook workbook = new HSSFWorkbook(f);
        List<Student> l = poiUtil.setWorkbook(workbook).get(1).convertWith(1, Student.class);
        l.forEach((x) -> {
            System.out.println(x.toString());
        });
    }
}

