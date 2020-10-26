package teprinciple.yang.list2excel;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class OfficeService {

    /**
     * 占位符
     */
    private static final Pattern SymbolPattern = Pattern.compile("\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);


    /**
     * 替换符号
     *
     * @param inputStream
     * @param symbolMap
     * @return
     * @throws IOException
     */
    public byte[] replaceSymbol(InputStream inputStream, Map<String, String> symbolMap) throws IOException {

        XWPFDocument doc = new XWPFDocument(inputStream);

        replaceSymbolInPara(doc, symbolMap);
        replaceInTable(doc, symbolMap);

        try (ByteArrayOutputStream os = new ByteArrayOutputStream()) {
            doc.write(os);
            return os.toByteArray();
        } finally {
            inputStream.close();
        }
    }


    /**
     * 替换Para中的Symbol
     *
     * @param doc
     * @param symbolMap
     */
    private void replaceSymbolInPara(XWPFDocument doc, Map<String, String> symbolMap) {
        XWPFParagraph para;
        Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
        while ( iterator.hasNext() ) {
            para = iterator.next();
            replaceInPara(para, symbolMap);
        }
    }

    /**
     * 替换正文
     */
    private void replaceInPara(XWPFParagraph para, Map<String, String> symbolMap) {

        List<XWPFRun> runs;
        if (symbolMatcher(para.getParagraphText()).find()) {
            String text = para.getParagraphText();
            Matcher matcher = SymbolPattern.matcher(text);
            while ( matcher.find() ) {
                String group = matcher.group(1);
                String symbol = symbolMap.get(group);
                if (symbol.isEmpty()) {
                    symbol = " ";
                }
                text = matcher.replaceFirst(symbol);
                matcher = SymbolPattern.matcher(text);
            }
            runs = para.getRuns();
            String fontFamily = runs.get(0).getFontFamily();
            int fontSize = runs.get(0).getFontSize();
            XWPFRun xwpfRun = para.insertNewRun(0);
            xwpfRun.setFontFamily(fontFamily);
            xwpfRun.setText(text);
            if (fontSize > 0) {
                xwpfRun.setFontSize(fontSize);
            }
            int max = runs.size();
            for (int i = 1; i < max; i++) {
                para.removeRun(1);
            }

        }
    }

    /**
     * 符号匹配器
     *
     * @param paragraphText
     * @return
     */
    private Matcher symbolMatcher(String paragraphText) {
        Matcher matcher = SymbolPattern.matcher(paragraphText);
        return matcher;
    }


    //替换表格
    private void replaceInTable(XWPFDocument doc, Map<String, String> symbolMap) {
        Iterator<XWPFTable> iterator = doc.getTablesIterator();
        XWPFTable table;
        List<XWPFTableRow> rows;
        List<XWPFTableCell> cells;
        List<XWPFParagraph> paras;
        while ( iterator.hasNext() ) {
            table = iterator.next();
            rows = table.getRows();
            for (XWPFTableRow row : rows) {
                cells = row.getTableCells();
                for (XWPFTableCell cell : cells) {
                    paras = cell.getParagraphs();
                    for (XWPFParagraph para : paras) {
                        replaceInPara(para, symbolMap);
                    }
                }
            }
        }
    }


}
