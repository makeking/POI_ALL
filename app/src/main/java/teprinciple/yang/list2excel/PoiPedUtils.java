package teprinciple.yang.list2excel;


import android.graphics.fonts.Font;

import org.apache.poi.hslf.record.Document;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.xwpf.usermodel.XWPFComment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

/**
 * 使用poi 导出pdf 的工具类
 */
public class PoiPedUtils {
    /**
     *  1. 生成doc 样式的内容
     * @param path
     */
    public static void exportDoc(String path) {


    }

//    public static void exportPef(String fileName) {
//        try {
//            // 第一步，实例化一个document对象
//            XWPFDocument document = new XWPFDocument();
//            // 第二步，设置要到出的路径
//            FileOutputStream out = null;
//            out = new FileOutputStream(fileName);
//
//            //如果是浏览器通过request请求需要在浏览器中输出则使用下面方式
//            //OutputStream out = response.getOutputStream();
//            // 第三步,设置字符
//            BaseFont bfChinese = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H", false);
//            Font fontZH = new Font(bfChinese, 12.0F, 0);
//            // 第四步，将pdf文件输出到磁盘
//            PdfWriter writer = PdfWriter.getInstance(document, out);
//            // 第五步，打开生成的pdf文件
//            document.open();
//            // 第六步,设置内容
//            String title = "标题";
//            document.add(new Paragraph(new Chunk(title, fontZH).setLocalDestination(title)));
//            document.add(new Paragraph("\n"));
//            // 创建table,注意这里的2是两列的意思,下面通过table.addCell添加的时候必须添加整行内容的所有列
//            PdfPTable table = new PdfPTable(2);
//            table.setWidthPercentage(100.0F);
//            table.setHeaderRows(1);
//            table.getDefaultCell().setHorizontalAlignment(1);
//            table.addCell(new Paragraph("序号", fontZH));
//            table.addCell(new Paragraph("结果", fontZH));
//            table.addCell(new Paragraph("1", fontZH));
//            table.addCell(new Paragraph("出来了", fontZH));
//
//            document.add(table);
//            document.add(new Paragraph("\n"));
//            // 第七步，关闭document
//            document.close();
//
//            System.out.println("导出pdf成功~");
//        } catch (FileNotFoundException e) {
//            e.printStackTrace();
//        }
//
//    }

} 
