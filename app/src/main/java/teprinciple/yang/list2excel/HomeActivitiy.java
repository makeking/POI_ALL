package teprinciple.yang.list2excel;

import android.Manifest;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.support.v4.app.ActivityCompat;
import android.support.v4.content.ContextCompat;
import android.support.v7.app.AppCompatActivity;
import android.view.View;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import teprinciple.yang.list2excel.exection.SheetExcetion;

import static android.content.pm.PackageManager.PERMISSION_GRANTED;

public class HomeActivitiy extends AppCompatActivity {
    private List<Student> students;
    private static String[] title = {"编号", "姓名", "性别", "年龄", "班级", "数学", "英语", "语文"};
    //private static String[] title = {"类别\n月份","1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月"};
    private ArrayList recordList;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_home_activitiy);
        // sd 卡授权
        initPro();

        initData();

    }

    private void initData() {
        //模拟数据集合
        students = new ArrayList<>();
        for (int i = 1; i <= 10; i++) {
            students.add(new Student("小红" + i, "女", 12, 1 + i, "一班", 85, 77, 98));
            students.add(new Student("小明" + i, "男", 14, 2 + i, "二班", 65, 57, 100));
        }
    }

    /**
     * 给应用授权的方法
     */
    private void initPro() {
        // 继续执行之后的操作
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.M) {
            if (ContextCompat.checkSelfPermission(this, Manifest.permission.WRITE_EXTERNAL_STORAGE) != PERMISSION_GRANTED) {
                ActivityCompat.requestPermissions(this, new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE}, 1);
            } else {
                Toast.makeText(this, "您已经申请了权限!", Toast.LENGTH_SHORT).show();
            }
        }
    }

    public void importXls(View view) throws Exception {
        File file = new File(getSDPath() + "/Record");
        makeDir(file);
        String fileName = file.getPath() + "/exportTest.xlsx";
        PoiExccelUtils.initXSSFSheet("成绩表");
        PoiExccelUtils.initExcecl(0, fileName, title, (short) 340);
        PoiExccelUtils.writeObjListToExcel(getRecordData(), fileName, 1);
        // 添加图片
        PoiExccelUtils.writePng(fileName, getSDPath() + "/11/Pictures/duola.png", new XSSFClientAnchor(0, 0, 0, 0, (short) 0, 22, (short) 4, 28), HSSFWorkbook.PICTURE_TYPE_PNG);
        // 添加折线图
        HashMap<String, CellRangeAddress> map = new HashMap<String, CellRangeAddress>();
        map.put("年龄", new CellRangeAddress(1, 20, 3, 3));
        PoiExccelUtils.writeCurve(fileName, PoiExccelUtils.getSheet().createDrawingPatriarch().createAnchor(0, 0, 0, 0, 8, 0, 20, 20),
                new CellRangeAddress(1, 20, 1, 1), map, 12, 14);
        PoiExccelUtils.mergeCells(0, 2, 0, 2, fileName);

    }


    /**
     * 将数据集合 转化成ArrayList<ArrayList<String>>
     *
     * @return
     */
    private ArrayList<ArrayList<Object>> getRecordData() {
        recordList = new ArrayList<>();
        for (int i = 0; i < students.size(); i++) {
            Student student = students.get(i);
            ArrayList<Object> beanList = new ArrayList<>();
            beanList.add(student.id);
            beanList.add(student.name);
            beanList.add(student.sex);
            beanList.add(student.age);
            beanList.add(student.classNo);
            beanList.add(student.math);
            beanList.add(student.english);
            beanList.add(student.chinese);
            recordList.add(beanList);
        }
        return recordList;
    }

    private String getSDPath() {
        File sdDir = null;
        boolean sdCardExist = Environment.getExternalStorageState().equals(
                android.os.Environment.MEDIA_MOUNTED);
        if (sdCardExist) {
            sdDir = Environment.getExternalStorageDirectory();
        }
        String dir = sdDir.toString();
        return dir;
    }

    public void makeDir(File dir) {
        if (!dir.getParentFile().exists()) {
            makeDir(dir.getParentFile());
        }
        dir.mkdir();
    }

    /**
     * 导出doc 文本
     *
     * @param view
     */
    public void exportDoc(View view) {
    }

    public void exportPdf(View view) {

    }
}