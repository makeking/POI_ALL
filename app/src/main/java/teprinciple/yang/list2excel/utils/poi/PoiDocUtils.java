//package teprinciple.yang.list2excel.utils.poi;
//
//import org.apache.poi.xwpf.usermodel.XWPFDocument;
//import org.apache.poi.xwpf.usermodel.XWPFParagraph;
//import org.apache.poi.xwpf.usermodel.XWPFRun;
//import java.io.File;
//import java.io.FileOutputStream;
//
///**
// * 初始化 doc 文档
// */
//public class PoiDocUtils {
//    /**
//     *      * 1. 普通文档
//     *      * 2. 表格
//     *      * 3. 折线图
//     */
//
//    /**
//     * 将一串文字设置到文件中
//     *
//     * @param data
//     * @param fileName
//     */
//
//    public static void writeFontToDoc(String data, String fileName) throws Exception {
//        // 像文件中写入数据
//        XWPFDocument doc = new XWPFDocument();// 创建Word文件
//        XWPFParagraph p2 = doc.createParagraph();//创建段落文本
//        XWPFRun r2 = p2.createRun();
//        r2.setText(data);
//        writeFile(doc, fileName);
//    }
//
//    public static void writerTitleToDoc(String data, String fileName) throws Exception {
//        // 像文件中写入数据
//        XWPFDocument doc = new XWPFDocument();// 创建Word文件
//        writeFile(doc, fileName);
//    }
//
//    public static  void writeTabToDoc(){
//
//    }
//
//    private static void writeFile(XWPFDocument doc, String fileName) throws Exception {
//        File file = new File(fileName);
//        if (!file.exists()) {
//            file.createNewFile();//创建文件夹
//        }
//        FileOutputStream out = new FileOutputStream(file);
//        doc.write(out);
//    }
//
//
//
//}
