package teprinciple.yang.list2excel.utils;

import org.apache.poi.POIXMLException;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.util.List;
import java.util.Map;

public class ReplaceUtils {
    public static void replaceParagraph(XWPFParagraph paragraph, Map<String, String> fieldsForReport) throws POIXMLException {
        String find, text, runsText;
        List<XWPFRun> runs;
        XWPFRun run, nextRun;
        for (String key : fieldsForReport.keySet()) {
            text = paragraph.getText();
            if (!text.contains("${"))
                return;
            find = key;
            if (!text.contains(find))
                continue;
            runs = paragraph.getRuns();
            for (int i = 0; i < runs.size(); i++) {
                run = runs.get(i);
                runsText = run.getText(0);
                if (runsText.contains("${") || (runsText.contains("$") && runs.get(i + 1).getText(0).substring(0, 1).equals("{"))) {
                    //As the next run may has a closed tag and an open tag at
                    //the same time, we have to be sure that our building string
                    //has a fully completed tags
                    while ( !openTagCountIsEqualCloseTagCount(runsText) ) {
                        nextRun = runs.get(i + 1);
                        runsText = runsText + nextRun.getText(0);
                        paragraph.removeRun(i + 1);
                    }
                    run.setText(runsText.contains(find) ?
                            runsText.replace(find, fieldsForReport.get(key)) :
                            runsText, 0);
                }
            }
        }
    }

    private static boolean openTagCountIsEqualCloseTagCount(String runText) {
        int openTagCount = runText.split("\\$\\{", -1).length == 0 ? -1 : runText.split("\\$\\{", -1).length - 1;
        int closeTagCount = runText.split("\\}", -1).length == 0 ? -2 :  runText.split("\\}", -1).length - 1;
        return openTagCount == closeTagCount;
    }
} 
