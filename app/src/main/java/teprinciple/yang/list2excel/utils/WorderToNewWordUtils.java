package teprinciple.yang.list2excel.utils;

import android.util.Log;
import android.widget.EditText;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class WorderToNewWordUtils {
    /**
     * 替换段落文本
     *
     * @param document docx解析对象
     * @param textMap  需要替换的信息集合
     */
    public static void changeText(XWPFDocument document, Map<String, String> textMap) {
        //获取段落集合
//        List<XWPFParagraph> paragraphs = document.getParagraphs();
        // 方法1
//        for (XWPFParagraph paragraph : paragraphs) {
//            //判断此段落时候需要进行替换
//            String text = paragraph.getText();
//            methodOne(textMap, paragraph, text);
//
//        }
        method4(textMap, document);

    }

    private static void method4(Map<String, String> textMap, XWPFDocument document) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (int j = 0; j < paragraphs.size(); j++) {
            XWPFParagraph paragraph = paragraphs.get(j);
            int start = -1, end = -1;
            boolean isStart = true;
            ArrayList<String> list = new ArrayList<>();
            StringBuilder data = new StringBuilder();
            do {
                List<XWPFRun> run = paragraph.getRuns();
                if (run.size() == 0) {
                    isStart = false;
                }
                if (isStart) {
                    for (int i = 0; i < run.size(); i++) {
                        if (i != run.size() - 1) {
                            data.append(run.get(i).toString());
                            if (run.get(i).toString().startsWith("$") && run.get(i).toString().endsWith("}")) {
                                start = end = i;
                                list.add(data.toString());
                                data.delete(0,data.length());
//                                break;
                            } else {
                                if ("$".equals(run.get(i).toString())) {
                                    start = i;
                                } else if ("}".equals(run.get(i).toString())) {
                                    end = i;
                                    list.add(data.toString());
                                    data.delete(0,data.length());
//                                    break;
                                } else if (run.get(i).toString().endsWith("}")) {
                                    end = i;
                                    list.add(data.toString());
                                    data.delete(0,data.length());
//                                    break;
                                } else if (run.get(i).toString().startsWith("$")) {
                                    start = i;
                                }
                            }
                        } else {
                            isStart = false;
                        }

                    }
                }
                if (start != -1 && end != -1 && isStart) {
                    for (int i = 0; i < run.size(); i++) {
                        if (start == end) {
                            run.get(start).setText(textMap.get(data.toString()), 0);
                            start = end = -1;
                            break;
                        } else {
                            if (i == start) {
                                run.get(start).setText(textMap.get(data.toString()), 0);
                            } else if (i <= end && i > start) {
                                run.get(i).setText("",0);
                            } else if (i > end) {
                                start = end = -1;
                                break;
                            }
                        }
                    }
                }
            }
            while ( isStart );
        }
    }

    private static void methodOne(Map<String, String> textMap, XWPFParagraph
            paragraph, String text) {
        if (checkText(text)) {
            List<XWPFRun> runs = paragraph.getRuns();
            for (XWPFRun run : runs) {
                run.setText(changeValue(run.toString(), textMap), 0);
            }
        }
    }


    /**
     * 替换表格对象方法
     *
     * @param document  docx解析对象
     * @param textMap   需要替换的信息集合
     * @param tableList 需要插入的表格信息集合
     */
    public static void changeTable(XWPFDocument document, Map<String, String> textMap,
                                   List<String[]> tableList) {
        //获取表格对象集合
        List<XWPFTable> tables = document.getTables();
        for (int i = 0; i < tables.size(); i++) {
            //只处理行数大于等于2的表格，且不循环表头
            XWPFTable table = tables.get(i);
            if (table.getRows().size() > 1) {
                //判断表格是需要替换还是需要插入，判断逻辑有$为替换，表格无$为插入
                if (checkText(table.getText())) {
                    List<XWPFTableRow> rows = table.getRows();
                    //遍历表格,并替换模板
                    eachTable(rows, textMap);
                } else {
//                  System.out.println("插入"+table.getText());
                    insertTable(table, tableList);
                }
            }
        }
    }

    /**
     * 遍历表格
     *
     * @param rows    表格行对象
     * @param textMap 需要替换的信息集合
     */
    public static void eachTable(List<XWPFTableRow> rows, Map<String, String> textMap) {
        for (XWPFTableRow row : rows) {
            List<XWPFTableCell> cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                //判断单元格是否需要替换
                if (checkText(cell.getText())) {
                    List<XWPFParagraph> paragraphs = cell.getParagraphs();
                    for (XWPFParagraph paragraph : paragraphs) {
                        List<XWPFRun> runs = paragraph.getRuns();
                        for (XWPFRun run : runs) {
                            run.setText(changeValue(run.toString(), textMap), 0);
                        }
                    }
                }
            }
        }
    }

    /**
     * 判断文本中时候包含$
     *
     * @param text 文本
     * @return 包含返回true, 不包含返回false
     */
    public static boolean checkText(String text) {
        boolean check = false;
        if (text.indexOf("$") != -1) {
            check = true;
        }
        return check;

    }

    /**
     * 为表格插入数据，行数不够添加新行
     *
     * @param table     需要插入数据的表格
     * @param tableList 插入数据集合
     */
    public static void insertTable(XWPFTable table, List<String[]> tableList) {
        //创建行,根据需要插入的数据添加新行，不处理表头
        for (int i = 1; i < tableList.size(); i++) {
            XWPFTableRow row = table.createRow();
        }
        //遍历表格插入数据
        List<XWPFTableRow> rows = table.getRows();
        for (int i = 1; i < rows.size(); i++) {
            XWPFTableRow newRow = table.getRow(i);
            List<XWPFTableCell> cells = newRow.getTableCells();
            for (int j = 0; j < cells.size(); j++) {
                XWPFTableCell cell = cells.get(j);
                cell.setText(tableList.get(i - 1)[j]);
            }
        }

    }

    /**
     * 匹配传入信息集合与模板
     *
     * @param value   模板需要替换的区域
     * @param textMap 传入信息集合
     * @return 模板需要替换区域信息集合对应值
     */
    public static String changeValue(String value, Map<String, String> textMap) {
        Set<Map.Entry<String, String>> textSets = textMap.entrySet();
        for (Map.Entry<String, String> textSet : textSets) {
            //匹配模板与替换值 格式${key}
            String key = textSet.getKey();
            if (value.contains(key)) {
                value = textSet.getValue();
            }
        }
        //模板未匹配到区域替换为空
        if (checkText(value)) {
            value = "";
        }
        return value;
    }


}
