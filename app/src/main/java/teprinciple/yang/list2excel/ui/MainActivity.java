package teprinciple.yang.list2excel.ui;

import android.Manifest;
import android.os.Build;
import android.os.Environment;
import android.support.v4.app.ActivityCompat;
import android.support.v4.content.ContextCompat;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.view.View;
import android.widget.Toast;


import java.io.File;
import java.util.ArrayList;
import java.util.List;

import teprinciple.yang.list2excel.utils.jxl.ExcelUtils;
import teprinciple.yang.list2excel.R;
import teprinciple.yang.list2excel.bean.Student;

import static android.content.pm.PackageManager.PERMISSION_GRANTED;

public class MainActivity extends AppCompatActivity {
    private ArrayList<ArrayList<Object>> recordList;
    private List<Student> students;
    private static String[] title = {"编号", "姓名", "性别", "年龄", "班级", "数学", "英语", "语文"};
    private File file;
    private String fileName;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        // sd 卡授权
        initPro();
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

    /**
     * 导出excel
     *
     * @param view
     */
    public void exportExcel(View view) {
        file = new File(getSDPath() + "/Record");
        makeDir(file);
        ExcelUtils.initExcel(file.getPath() + "/成绩表.xls", title);
        fileName = getSDPath() + "/Record/成绩表.xls";
        ExcelUtils.writeObjListToExcel(getRecordData(), fileName, this);

//        ExcelUtils.writePng(file.getPath() + "/成绩表.xls", getSDPath() + "/11/Pictures/duola.png", 1, recordList.size() + 2);

        // 生成图表


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
            ArrayList<Object> beanList = new ArrayList<Object>();
            beanList.add(student.getId());
            beanList.add(student.getName());
            beanList.add(student.getSex());
            beanList.add(student.getAge());
            beanList.add(student.getClassNo());
            beanList.add(student.getMath());
            beanList.add(student.getEnglish());
            beanList.add(student.getChinese());
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

    public void demos(View view) {

    }
}
