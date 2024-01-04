package com.lh.word;

import com.lh.word.form.BarChartForm;
import com.lh.word.form.ScatterChartForm;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.chart.BarGrouping;
import org.apache.poi.xddf.usermodel.chart.MarkerStyle;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Test {
    public static void main(String[] args) {
//        try (XWPFDocument document = new XWPFDocument(new FileInputStream("D:\\FreeMarker.docx"))) {
//            WordUtils wordUtils = new WordUtils();
//            Map<String, String> paragraphMap = new HashMap<>();
//            paragraphMap.put("number", "10000");
//            paragraphMap.put("date", "2020-03-25");
//            wordUtils.changeParagraphText(document, paragraphMap);
//
//            List<Map<String, String>> tableTextList = new ArrayList<>();
//            Map<String, String> tableMap = new HashMap<>();
//            tableMap.put("name", "赵云");
//            tableMap.put("sexual", "男");
//            tableMap.put("birthday", "2020-01-01");
//            tableMap.put("identify", "123456789");
//            tableMap.put("phone", "18377776666");
//            tableMap.put("address", "王者荣耀");
//            tableMap.put("domicile", "中国-腾讯");
//            tableMap.put("QQ", "是");
//            tableMap.put("chat", "是");
//            tableMap.put("blog", "是");
//            tableTextList.add(tableMap);
//            Map<String, String> tableMap2 = new HashMap<>();
//            tableMap2.put("spring", "sony的名称");
//            tableTextList.add(tableMap2);
//            wordUtils.changeTableText(document, tableTextList);
//
//            List<TableForm> list = new ArrayList<>();
//            TableForm tableForm = new TableForm();
//            tableForm.setStartLine(7);
//            tableForm.getData().add(new String[]{"露娜", "女", "野友", "666", "6660"});
//            tableForm.getData().add(new String[]{"鲁班", "男", "射友", "222", "2220"});
//            tableForm.getData().add(new String[]{"程咬金", "男", "肉友", "999", "9990"});
//            tableForm.getData().add(new String[]{"太乙真人", "男", "辅友", "111", "1110"});
//            tableForm.getData().add(new String[]{"貂蝉", "女", "法友", "888", "8880"});
//            list.add(tableForm);
//            TableForm tableForm2 = new TableForm();
//            tableForm2.setStartLine(1);
//            tableForm2.getData().add(new String[]{"18581588710", "蜘蛛侠", "100"});
//            tableForm2.getData().add(new String[]{"18581588710", "战神", "200"});
//            list.add(tableForm2);
//            wordUtils.copyHeaderInsertText(document,list);
//            try (FileOutputStream fileOut = new FileOutputStream("CreateWordXDDFChart.docx")) {
//                document.write(fileOut);
//            }
//        } catch (Exception e) {
//
//        }


//        try (XWPFDocument document = new XWPFDocument()) {
//            WordUtils wordUtils = new WordUtils();
//            XWPFChart chart = wordUtils.getChart(document, null, null);
//            PieChartForm pieChartForm = new PieChartForm();
//            pieChartForm.setTitle("标题");
//            pieChartForm.setBottomData(new String[]{"俄罗斯", "加拿大", "美国", "中国", "巴西", "澳大利亚", "印度"});
//            pieChartForm.setLeftData(new Integer[]{17098242, 9984670, 9826675, 9596961, 8514877, 7741220, 3287263});
//            wordUtils.createPieChart(chart, pieChartForm);
//            try (FileOutputStream fileOut = new FileOutputStream("CreateWordXDDFChart.docx")) {
//                document.write(fileOut);
//            }
//        } catch (Exception e) {
//
//        }

        /**
         * 散点图
         */
        try (XWPFDocument document = new XWPFDocument()) {
            WordUtils wordUtils = new WordUtils();
            XWPFChart chart = wordUtils.getChart(document, null, null);
            ScatterChartForm scatterChartForm = new ScatterChartForm();
            scatterChartForm.setTitle("测试");
            scatterChartForm.setBottomTitle("X轴");
            scatterChartForm.setLeftTitle("Y轴");
            scatterChartForm.setStyle(MarkerStyle.CIRCLE);
            scatterChartForm.setMarkerSize((short) 10);
            scatterChartForm.setVaryColors(false);

            ScatterChartForm.AreaData areaData = new ScatterChartForm.AreaData();
            areaData.setBottomData(new Integer[]{1, 2, 3, 4, 5, 8, 7});
            areaData.setLeftData(new Integer[]{5, 5, 5, 4, 5, 6, 7});
            areaData.setTitle("测试1");
            scatterChartForm.getLists().add(areaData);

            ScatterChartForm.AreaData areaData2 = new ScatterChartForm.AreaData();
            areaData2.setBottomData(new Integer[]{6,9});
            areaData2.setLeftData(new Integer[]{1,9});
            areaData2.setXddfColor(XDDFColor.from(new byte[]{(byte)0xFF, (byte)0xE1, (byte)0xFF}));
            areaData2.setTitle("测试2");
            scatterChartForm.getLists().add(areaData2);
            wordUtils.createScatterChart(chart, scatterChartForm);

            try (FileOutputStream fileOut = new FileOutputStream("D://CreateWordXDDFChart.docx")) {
                document.write(fileOut);
            }
        } catch (Exception e) {

        }

//        try (XWPFDocument document = new XWPFDocument()) {
//            WordUtils wordUtils = new WordUtils();
//            XWPFChart chart = wordUtils.getChart(document, null, null);
//            LineChartForm lineChartForm = new LineChartForm();
//            lineChartForm.setTitle("测试");
//            lineChartForm.setBottomTitle("X轴");
//            lineChartForm.setLeftTitle("Y轴");
//            lineChartForm.setStyle(MarkerStyle.STAR);
//            lineChartForm.setMarkerSize((short) 6);
//            lineChartForm.setSmooth(false);
//            lineChartForm.setVaryColors(false);
//            lineChartForm.setBottomData(new String[] {"俄罗斯","加拿大","美国","中国","巴西","澳大利亚","印度"});
//            lineChartForm.setLeftData(new Integer[] {17098242,9984670,9826675,9596961,8514877,7741220,3287263});
//            wordUtils.createLineChart(chart, lineChartForm);
//            try (FileOutputStream fileOut = new FileOutputStream("CreateWordXDDFChart.docx")) {
//                document.write(fileOut);
//            }
//        } catch (Exception e) {
//
//        }


//        try (XWPFDocument document = new XWPFDocument()) {
//            WordUtils wordUtils = new WordUtils();
//            XWPFChart chart = wordUtils.getChart(document, null, null);
//            String[] categories = new String[]{"Lang 1", "Lang 2", "Lang 3"};
//            Double[] valuesA = new Double[]{10d, 20d, 30d};
//            Double[] valuesB = new Double[]{15d, 25d, 35d};
//            Double[] valuesC = new Double[]{10d, 8d, 20d};
//            List<Double[]> list = new ArrayList<>();
//            list.add(valuesA);
//            list.add(valuesB);
//            list.add(valuesC);
//            BarChartForm barChartForm = new BarChartForm();
//            barChartForm.setTitle("测试");
//            barChartForm.setCategories(categories);
//            barChartForm.setTableData(list);
//            barChartForm.setColorTitles(Arrays.asList("a", "b", "c"));
//            barChartForm.setGrouping(BarGrouping.STACKED);
//            barChartForm.setNewOverlap((byte) 100);
//
//            BarChartForm.ColorCheck colorCheck = new BarChartForm.ColorCheck();
//            colorCheck.setXddfColor(XDDFColor.from(new byte[]{(byte) 0xFF, (byte) 0x33, (byte) 0x00}));
//            colorCheck.setNum(0);
//            barChartForm.getList().add(colorCheck);
//
//            BarChartForm.ColorCheck colorCheck2 = new BarChartForm.ColorCheck();
//            colorCheck2.setXddfColor(XDDFColor.from(new byte[]{(byte) 0x91, (byte) 0x2C, (byte) 0xEE}));
//            colorCheck2.setNum(1);
//            barChartForm.getList().add(colorCheck2);
//
//            BarChartForm.ColorCheck colorCheck3 = new BarChartForm.ColorCheck();
//            colorCheck3.setXddfColor(XDDFColor.from(new byte[]{(byte) 0x00, (byte) 0x00, (byte) 0x80}));
//            colorCheck3.setNum(2);
//            barChartForm.getList().add(colorCheck3);
//
//
//            wordUtils.createBarChart(chart, barChartForm);
//            try (FileOutputStream fileOut = new FileOutputStream("CreateWordXDDFChart.docx")) {
//                document.write(fileOut);
//            }
//        } catch (Exception e) {
//
//        }
    }
}
