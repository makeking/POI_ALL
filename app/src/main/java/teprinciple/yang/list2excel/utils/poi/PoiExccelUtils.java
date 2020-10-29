package teprinciple.yang.list2excel.utils.poi;

import android.support.annotation.NonNull;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.charts.AxisCrosses;
import org.apache.poi.ss.usermodel.charts.AxisPosition;
import org.apache.poi.ss.usermodel.charts.ChartDataSource;
import org.apache.poi.ss.usermodel.charts.DataSources;
import org.apache.poi.ss.usermodel.charts.LegendPosition;
import org.apache.poi.ss.usermodel.charts.LineChartData;
import org.apache.poi.ss.usermodel.charts.LineChartSeries;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.charts.XSSFCategoryAxis;
import org.apache.poi.xssf.usermodel.charts.XSSFChartLegend;
import org.apache.poi.xssf.usermodel.charts.XSSFValueAxis;
import org.w3c.dom.Document;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Map;
import java.util.Objects;

import teprinciple.yang.list2excel.exection.NumberExcetion;
import teprinciple.yang.list2excel.exection.SheetExcetion;

/**
 * 使用poi 导出excel 的工具类
 */
public class PoiExccelUtils {

    private static XSSFSheet sheet;
    //    private static HSSFSheet sheet;
    private static XSSFWorkbook hssfWorkbook;

    /**
     * 初始化 XSSFSheet，如果工具类没有调用这个方法，就报出异常
     *
     * @param sheetName
     * @return
     */
    public static XSSFSheet initXSSFSheet(String sheetName) {
        if (hssfWorkbook == null || sheet == null) {
            hssfWorkbook = new XSSFWorkbook();
            sheet = hssfWorkbook.createSheet(sheetName);
        }
        return sheet;
    }

    /**
     * 创建标题行
     */
    public static void initExcecl(@NonNull int rowNum, @NonNull String path, String[] title, short... heights) throws Exception {
        short height;
        if (hssfWorkbook == null || sheet == null) {
            throw new SheetExcetion("can not find XSSFSheet or XSSFWorkbook ,please use initXSSFSheet method !!! ");
        }
        //创建行,行号作为参数传递给createRow()方法,第一行从0开始计算
        XSSFRow row = sheet.createRow(rowNum);
        //创建单元格,row已经确定了行号,列号作为参数传递给createCell(),第一列从0开始计算
        for (int i = 0; i < title.length; i++) {
            XSSFCell cell = row.createCell(i);
            //设置单元格的值,即C1的值(第一行,第三列)
            cell.setCellValue(title[i]);
        }
        if (heights.length == 0) {
            height = 340;
        } else if (heights.length == 1) {
            height = heights[0];
        } else {
            throw new NumberExcetion(" You must Incoming one or null length , Otherwise, this exception will be reported !!! ");
        }
        row.setHeight(height);
        // 字体加粗， 文字居中
        //输出到磁盘中
        FileOutputStream fos;
        try {
            fos = new FileOutputStream(new File(path));
            hssfWorkbook.write(fos);
//            hssfWorkbook.close();
            fos.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 传入一个数组，创建一个表格
     *
     * @param objList
     * @param fileName
     * @param rowNum
     * @param <T>
     * @throws Exception
     */
    public static <T> void writeObjListToExcel(@NonNull ArrayList<T> objList, @NonNull String fileName, @NonNull int rowNum) throws Exception {
        if (hssfWorkbook == null || sheet == null) {
            throw new SheetExcetion("can not find XSSFSheet or XSSFWorkbook ,please use initXSSFSheet method !!! ");
        }
        if (rowNum <= 0) {
            throw new NumberExcetion("You must Incoming number is above zero ,  Otherwise, this exception will be reported !!!");
        }
        if (!new File(fileName).exists()) {
            throw new FileNotFoundException("ou must  Incoming a exit file, Otherwise, this exception will be reported !!!");
        }
        //创建行,行号作为参数传递给createRow()方法,第一行从0开始计算
        for (int i = rowNum; i < objList.size() + rowNum; i++) {
            ArrayList<Object> list = (ArrayList<Object>) objList.get(i - rowNum);
            XSSFRow row = sheet.createRow(i);
            for (int j = 0; j < list.size(); j++) {
                XSSFCell cell = row.createCell(j);
                //设置单元格的值,即C1的值(第一行,第三列)
                // 判断数据的类型
                Object data = list.get(j);
                if (data instanceof Integer) {
                    int value = ((Integer) data).intValue();
                    cell.setCellValue(value);
                } else if (data instanceof String) {
                    String s = (String) data;
                    cell.setCellValue(s);
                } else if (data instanceof Double) {
                    double d = ((Double) data).doubleValue();
                    cell.setCellValue(d);
                } else if (data instanceof Float) {
                    float f = ((Float) data).floatValue();
                    cell.setCellValue(f);
                } else if (data instanceof Long) {
                    long l = ((Long) data).longValue();
                    cell.setCellValue(l);
                } else if (data instanceof Boolean) {
                    boolean b = ((Boolean) data).booleanValue();
                    cell.setCellValue(b);
                }

            }
        }
        //输出到磁盘中
        FileOutputStream fos;
        try {
            fos = new FileOutputStream(new File(fileName));
            hssfWorkbook.write(fos);
            fos.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 传入一张图片，将图片导入到excel 中
     *
     * @param fileName
     * @param pngPath
     * @param anchor
     * @param format
     * @throws Exception
     */
    public static void writePng(@NonNull String fileName, @NonNull String pngPath, @NonNull XSSFClientAnchor anchor, @NonNull int format) throws Exception {
        if (hssfWorkbook == null || sheet == null) {
            throw new SheetExcetion("can not find XSSFSheet or XSSFWorkbook ,please use initXSSFSheet method !!! ");
        }
        if (!new File(fileName).exists() || !new File(pngPath).exists()) {
            throw new FileNotFoundException("You must  Incoming a exit file, Otherwise, this exception will be reported !!!");
        }
        FileOutputStream fileOut = null;
        try {
            ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
            
            // todo java 的方式
//            bufferImg = ImageIO.read(new File(pngPath));
//            ImageIO.write(bufferImg, "PNG", byteArrayOut);
            //todo Android 的方式
            FileInputStream is = new FileInputStream(new File(pngPath));
            byte[] b = new byte[1024 * 2];
            int len;
            while ( (len = is.read(b, 0, b.length)) != -1 ) {
                byteArrayOut.write(b, 0, len);
                byteArrayOut.flush();
            }
            //画图的顶级管理器，一个sheet只能获取一个（一定要注意这点）
//            HSSFPatriarch
            XSSFDrawing patriarch = sheet.createDrawingPatriarch();
            //anchor主要用于设置图片的属性
            /*
                dx1 dy1 起始单元格中的x,y坐标.

                dx2 dy2 结束单元格中的x,y坐标

                col1,row1 指定起始的单元格，下标从0开始

                col2,row2 指定结束的单元格 ，下标从0开始
             */
//            HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 0, 0, (short) 21, 0, (short) 24, 2);
            // c1 表示的是右上角
            // c2 表示的是 5 表示距离左边
            anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
            //插入图片
            patriarch.createPicture(anchor, hssfWorkbook.addPicture(byteArrayOut.toByteArray(), format));
            fileOut = new FileOutputStream(fileName);
            // 写入excel文件
            hssfWorkbook.write(fileOut);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (fileOut != null) {
                try {
                    fileOut.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

    }

    public static XSSFSheet getSheet() {
        return sheet;
    }

    /**
     * 设置折线图
     *
     * @param fileName
     * @param <T>
     */
    public static <T> void writeCurve(@NonNull String fileName, @NonNull XSSFClientAnchor anchor, @NonNull CellRangeAddress xCellRangeAddress, @NonNull Map<String, CellRangeAddress> maps, int... leftSize) throws Exception {
        if (hssfWorkbook == null || sheet == null) {
            throw new SheetExcetion("can not find XSSFSheet or XSSFWorkbook ,please use initXSSFSheet method !!! ");
        }
        if (!new File(fileName).exists()) {
            throw new FileNotFoundException("You must Incoming a exit file, Otherwise, this exception will be reported !!!");
        }
        if (maps.size() <= 0) {
            throw new NumberExcetion("You must  Incoming one or more map @param(Map<String,CellRangeAddress>), Otherwise, this exception will be reported !!!");
        }
        if (leftSize == null || leftSize.length != 2) {
            throw new NumberExcetion("You must  Incoming zero or tow int number, Otherwise, this exception will be reported !!!");
        }
        //1. 设置表格
        try {
            XSSFDrawing drawing = sheet.createDrawingPatriarch();
//            // 设置画布在excel 表中的位置
//            /**
//             *  前四个参数 ：
//             *  后4个参数 ： x ，y 轴的开始和结束的位置
//             */
//            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 8, 0, 20, 20);
            //获取图表
            XSSFChart chart = drawing.createChart(anchor);
//                chart.setTitle("2020届年龄曲线图");
            // 图表图例
            XSSFChartLegend legend = chart.getOrCreateLegend();
            // 设置标题的位置
            legend.setPosition(LegendPosition.TOP);
            // 创建折线图 LineChartData
            LineChartData chartData = chart.getChartDataFactory().createLineChartData();
            //设置横坐标
            XSSFCategoryAxis bottomAxis = chart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
            bottomAxis.setCrosses(AxisCrosses.AUTO_ZERO);
            // 取消x轴的刻度
//                bottomAxis.setMajorTickMark(AxisTickMark.NONE);
            //设置纵坐标
            XSSFValueAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
            leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
            if (2 == leftSize.length) {
                leftAxis.setMinimum(leftSize[0]);
                leftAxis.setMaximum(leftSize[1]);
            } else {
                leftAxis.setMinimum(0);
                leftAxis.setMaximum(1);
            }

            //从Excel中获取数据
            // x 轴
//            ChartDataSource<String> xAxis = DataSources.fromStringCellRange(sheet, new CellRangeAddress(1, 20, 1, 1));
            ChartDataSource<String> xAxis = DataSources.fromStringCellRange(sheet, xCellRangeAddress);
            for (String key : maps.keySet()) {
                ChartDataSource<Number> dataAxis = DataSources.fromNumericCellRange(sheet, Objects.requireNonNull(maps.get(key)));
                //将获取到的数据填充到折线图
                LineChartSeries chartSeries = chartData.addSeries(xAxis, dataAxis);
                //设置要绘制的内容
                chartSeries.setTitle(key);
            }
            // y 轴
//            ChartDataSource<Number> dataAxis = DataSources.fromNumericCellRange(sheet, new CellRangeAddress(1, 20, 3, 3));
//            //将获取到的数据填充到折线图
//            LineChartSeries chartSeries = chartData.addSeries(xAxis, dataAxis);
//            //设置要绘制的内容
//            chartSeries.setTitle(sheet.getRow(0).getCell(3).getStringCellValue());
            //开始绘制折线图
            chart.plot(chartData, bottomAxis, leftAxis);

        } catch (Exception exction) {
        }
        /* Finally define FileOutputStream and write chart information */
        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream(fileName);
            hssfWorkbook.write(fileOut);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (fileOut != null) {
                    fileOut.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 合并单元格的操作
     */

    public static void mergeCells(@NonNull int startRow, @NonNull int endRow, @NonNull int startCol, @NonNull int endCol, @NonNull String fileName) throws Exception {
        if (hssfWorkbook == null || sheet == null) {
            throw new SheetExcetion("can not find XSSFSheet or XSSFWorkbook ,please use initXSSFSheet method !!! ");
        }
        if (!new File(fileName).exists()) {
            throw new FileNotFoundException("You must  Incoming a exit file, Otherwise, this exception will be reported !!!");
        }

        try {
            CellRangeAddress region = new CellRangeAddress(startRow, endRow, startCol, endCol);
            sheet.addMergedRegion(region);
            File file = new File(fileName);
            FileOutputStream fout;
            fout = new FileOutputStream(file);
            hssfWorkbook.write(fout);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
