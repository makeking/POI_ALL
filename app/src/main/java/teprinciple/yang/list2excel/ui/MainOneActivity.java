package teprinciple.yang.list2excel.ui;

import android.Manifest;
import android.os.Build;
import android.support.v4.app.ActivityCompat;
import android.support.v4.content.ContextCompat;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.view.View;
import android.widget.Toast;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import teprinciple.yang.list2excel.R;
import teprinciple.yang.list2excel.utils.WorderToNewWordUtils;

import static android.content.pm.PackageManager.PERMISSION_GRANTED;

/**
 * 根据模板生成pdf/excel /doc
 */
public class MainOneActivity extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main_one);
        // sd 卡授权
        initPro();
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
     * 根据模板导出 Doc
     *
     * @param view
     */
    public void exportDoc(View view) {


        new Thread(new Runnable() {
            @Override
            public void run() {
                // 根据文件获取新的文件
                XWPFDocument docx = null;
                try {
                    File file = new File("mnt/sdcard/11/数据.docx");
                    FileInputStream inputStream = new FileInputStream(file);
                    docx = new XWPFDocument(inputStream);
                    setParagraph(docx);
//                    setTable(docx);
                    FileOutputStream fos = new FileOutputStream(new File("mnt/sdcard/11/新数据.docx"));
                    docx.write(fos);
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }).start();


    }


    private void setTable(XWPFDocument docx) {

    }

    private void setParagraph(XWPFDocument docx) {
        Map<String, String> params = new HashMap<>();
        params.put("${bigtitle}", "测试");
        params.put("${smallTitle}", "文字");
        params.put("${smallTitle2}", "表格");
        params.put("${smallTitle3}", "图表");
        params.put("${smallTitle4}", "图片");
        WorderToNewWordUtils.changeText(docx, params);
    }

    /**
     * 根据模板导出 Excel
     *
     * @param view
     */
    public void exportExcel(View view) {
    }

    /**
     * 根据模板导出pdf
     *
     * @param view
     */
    public void exportPdf(View view) {
    }
}