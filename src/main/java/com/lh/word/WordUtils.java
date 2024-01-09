package com.lh.word;

import com.lh.word.form.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.XDDFShapeProperties;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * Copyright (C), 2006-2010, ChengDu ybya info. Co., Ltd.
 * FileName: WordUtils.java
 *
 * @author lh
 * @version 1.0.0
 * @Date 2021/02/02 16:04
 */
public class WordUtils {

    /**
     * 获取图表对象
     *
     * @param document word对象
     * @param width    默认15
     * @param height   默认10
     * @return
     */
    public XWPFChart getChart(XWPFDocument document, Integer width, Integer height) throws IOException, InvalidFormatException {
        if (width == null) {
            width = 15;
        }
        if (height == null) {
            height = 10;
        }
        return document.createChart(width * Units.EMU_PER_CENTIMETER, height * Units.EMU_PER_CENTIMETER);
    }

    /**
     * 创建普通柱状图-簇状柱状图-堆叠柱状图
     *
     * @param chart        图表对象
     * @param barChartForm 数据对象
     */
    public void createBarChart(XWPFChart chart, BarChartForm barChartForm) throws Exception {
        String[] categories = barChartForm.getCategories();
        List<Double[]> tableData = barChartForm.getTableData();
        List<String> colorTitles = barChartForm.getColorTitles();
        String title = barChartForm.getTitle();
        if (colorTitles.size() != tableData.size()) {
            throw new Exception("颜色标题个数,必须和数组个数相同");
        }
        for (Double[] tableDatum : tableData) {
            if (tableDatum.length != categories.length) {
                throw new Exception("每个数组的元素个数,必须和");
            }
        }
        // 设置标题
        chart.setTitleText(title);
        //标题覆盖
        chart.setTitleOverlay(false);

        // 处理对应的数据
        int numOfPoints = categories.length;
        String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
        XDDFDataSource<String> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);
        List<XDDFChartData.Series> seriesList = new ArrayList<>();

        // 创建一些轴
        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle(barChartForm.getBottomTitle());
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle(barChartForm.getBottomTitle());
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
        leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
        // 创建柱状图的类型
        XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
        // 为图表添加数据
        for (int i = 0; i < tableData.size(); i++) {
            XDDFChartData.Series series = data.addSeries(categoriesData, XDDFDataSourcesFactory.fromArray(
                    tableData.get(i), chart.formatRange(new CellRangeAddress(1, numOfPoints, i, i))));
            seriesList.add(series);
        }
        for (int i = 0; i < seriesList.size(); i++) {
            seriesList.get(i).setTitle(colorTitles.get(i), setTitleInDataSheet(chart, colorTitles.get(i), 1));
        }
        // 指定为簇状柱状图
        if (tableData.size() > 1) {
            ((XDDFBarChartData) data).setBarGrouping(barChartForm.getGrouping());
            chart.getCTChart().getPlotArea().getBarChartArray(0).addNewOverlap().setVal(barChartForm.getNewOverlap());
        }

        // 指定系列颜色
        for (BarChartForm.ColorCheck colorCheck : barChartForm.getList()) {
            XDDFSolidFillProperties fillMarker = new XDDFSolidFillProperties(colorCheck.getXddfColor());
            XDDFShapeProperties propertiesMarker = new XDDFShapeProperties();
            // 给对象填充颜色属性
            propertiesMarker.setFillProperties(fillMarker);
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(colorCheck.getNum()).addNewSpPr().set(propertiesMarker.getXmlObject());
        }

        ((XDDFBarChartData) data).setBarDirection(BarDirection.COL);
        // 设置多个柱子之间的间隔
        // 绘制图形数据
        chart.plot(data);
        // create legend
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.LEFT);
        legend.setOverlay(false);


    }


    /**
     * 创建折线图
     *
     * @param chart         图表对象
     * @param lineChartForm 数据对象
     */
    public void createLineChart(XWPFChart chart, LineChartForm lineChartForm) {
        // 标题
        chart.setTitleText(lineChartForm.getTitle());
        //标题覆盖
        chart.setTitleOverlay(false);
        //图例位置
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP);
        //分类轴标(X轴),标题位置
        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle(lineChartForm.getBottomTitle());
        //值(Y轴)轴,标题位置
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle(lineChartForm.getLeftTitle());
        // 处理数据
        XDDFCategoryDataSource bottomDataSource = XDDFDataSourcesFactory.fromArray(lineChartForm.getBottomData());
        XDDFNumericalDataSource<Integer> leftDataSource = XDDFDataSourcesFactory.fromArray(lineChartForm.getLeftData());

        // 生成数据
        XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);

        // 不自动生成颜色
        data.setVaryColors(lineChartForm.getVaryColors());

        //图表加载数据，折线1
        XDDFLineChartData.Series series = (XDDFLineChartData.Series) data.addSeries(bottomDataSource, leftDataSource);

        //是否弯曲
        series.setSmooth(lineChartForm.getSmooth());

        //设置标记样式
        series.setMarkerStyle(lineChartForm.getStyle());

        //绘制
        chart.plot(data);
    }

    /**
     * 创建散点图
     *
     * @param chart            图表对象
     * @param scatterChartForm 数据对象
     */
    public void createScatterChart(XWPFChart chart, ScatterChartForm scatterChartForm) {
        // 标题
        chart.setTitleText(scatterChartForm.getTitle());
        //标题覆盖
        chart.setTitleOverlay(false);
        //图例位置
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP);
        //分类轴标(X轴),标题位置
        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle(scatterChartForm.getBottomTitle());
        //值(Y轴)轴,标题位置
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle(scatterChartForm.getLeftTitle());

        /**
         * 网格线
         */
        if (scatterChartForm.getIsShowYGrid()){
            leftAxis.getOrAddMajorGridProperties().setLineProperties(scatterChartForm.line());
        }

        if (scatterChartForm.getIsShowXGrid()){
            bottomAxis.getOrAddMajorGridProperties().setLineProperties(scatterChartForm.line());
        }

        //fix poi Bug see
        /**
         * fix poi Bug
         * 解决新版Microsoft Office打不开docx文件 wps可以打开
         * https://stackoverflow.com/questions/77131316/create-a-chart-in-a-powerpoint-using-apache-poi-and-java-from-scratch
         */
        if (bottomAxis.hasNumberFormat()) bottomAxis.setNumberFormat("@");
        if (leftAxis.hasNumberFormat()) leftAxis.setNumberFormat("#,##0.00");

        XDDFScatterChartData data = null;
        for (int i = 0; i < scatterChartForm.getLists().size(); i++) {
            // 处理数据
            XDDFNumericalDataSource bottomDataSource = XDDFDataSourcesFactory.fromArray(scatterChartForm.getLists().get(i).getBottomData());
            XDDFNumericalDataSource<Integer> leftDataSource = XDDFDataSourcesFactory.fromArray(scatterChartForm.getLists().get(i).getLeftData());
            // 生成数据
            if (data == null) {
                data = (XDDFScatterChartData) chart.createData(ChartTypes.SCATTER, bottomAxis, leftAxis);
                // 是否自动生成颜色
                data.setVaryColors(false);
            }

            //图表加载数据，折线1
            XDDFScatterChartData.Series series = (XDDFScatterChartData.Series) data.addSeries(bottomDataSource, leftDataSource);
            //设置标记样式
            series.setMarkerStyle(scatterChartForm.getStyle());
            series.setMarkerSize(scatterChartForm.getMarkerSize());
            // 设置系列标题
            series.setTitle(scatterChartForm.getLists().get(i).getTitle(), null);
            // 去除连接线
            chart.getCTChart().getPlotArea().getScatterChartArray(0).getSerArray(i).addNewSpPr().addNewLn().addNewNoFill();
            if (scatterChartForm.getLists().get(i).getXddfColor() != null) {
                // 创建一个设置对象
                XDDFSolidFillProperties fillMarker = new XDDFSolidFillProperties(scatterChartForm.getLists().get(i).getXddfColor());
                XDDFShapeProperties propertiesMarker = new XDDFShapeProperties();
                // 给对象填充颜色属性
                propertiesMarker.setFillProperties(fillMarker);
                // 修改系列颜色
                chart.getCTChart().getPlotArea().getScatterChartArray(0).getSerArray(i).getMarker()
                        .addNewSpPr().set(propertiesMarker.getXmlObject());
            }
        }


        if(!scatterChartForm.getIsShowLegend()){
            XmlCursor xmlCursor = chart.getCTChart().newCursor();
            removeTag(xmlCursor,"legend");
        }


        //绘制
        chart.plot(data);


    }


    /**
     * 删除指定的标签
     * @param cursor
     * @param tag
     */
    public static void removeTag(XmlCursor cursor, String tag){
        while(cursor.hasNextToken()){
            if(cursor.toNextToken().isStart()){
                //System.out.println(cursor.getName().getLocalPart());
                if(cursor.getName().getLocalPart().equals(tag)){
                    cursor.removeXml();
                }
            }
        }

        cursor.dispose();
    }

    /**
     * 创建饼状图
     *
     * @param chart        图表对象
     * @param pieChartForm 数据对象
     */
    public void createPieChart(XWPFChart chart, PieChartForm pieChartForm) {
        // 标题
        chart.setTitleText(pieChartForm.getTitle());
        //标题覆盖
        chart.setTitleOverlay(false);
        //图例位置
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP);
        // 处理数据
        XDDFCategoryDataSource bottomDataSource = XDDFDataSourcesFactory.fromArray(pieChartForm.getBottomData());
        XDDFNumericalDataSource<Integer> leftDataSource = XDDFDataSourcesFactory.fromArray(pieChartForm.getLeftData());

        // 生成数据
        XDDFChartData data = chart.createData(ChartTypes.PIE, null, null);
        // 自动生成颜色
        data.setVaryColors(false);

        //图表加载数据
        XDDFChartData.Series series = data.addSeries(bottomDataSource, leftDataSource);

        //绘制
        chart.plot(data);
    }

    /**
     * 添加word中的标记数据 标记方式为 ${text}
     *
     * @param document word对象
     * @param textMap  需要替换的信息集合
     */
    public void changeParagraphText(XWPFDocument document, Map<String, String> textMap) {
        //获取段落集合
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (XWPFParagraph paragraph : paragraphs) {
            //判断此段落时候需要进行替换
            String text = paragraph.getText();
            if (checkText(text)) {
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    //替换模板原来位置
                    run.setText(changeValue(run.toString(), textMap), 0);
                }
            }
        }
    }

    /**
     * 替换表格中标记的数据 标记方式为 ${text}
     * 这里有个奇怪的问题 输入${}符号的时候需要把输入法切换到中文
     * ${}中间不能用数字,不能有下划线
     *
     * @param document word对象
     * @param textMap  需要替换的信息集合
     */
    public void changeTableText(XWPFDocument document, List<Map<String, String>> tableTextList) {
        //获取表格对象集合
        List<XWPFTable> tables = document.getTables();
        for (int i = 0; i < tables.size(); i++) {
            Map<String, String> textMap = tableTextList.get(i);
            //只处理行数大于等于2的表格
            XWPFTable table = tables.get(i);
            if (table.getRows().size() > 1) {
                //判断表格是需要替换还是需要插入，判断逻辑有$为替换，表格无$为插入
                if (checkText(table.getText())) {
                    List<XWPFTableRow> rows = table.getRows();
                    //遍历表格,并替换模板
                    eachTable(rows, textMap);
                }
            }
        }
    }

    /**
     * 复制表头,插入行数据,这里的样式和表头一样
     *
     * @param document word对象
     * @param list     集合个数和word中的表格个数必须相同
     */
    public void copyHeaderInsertText(XWPFDocument document, List<TableForm> list) {
        //获取表格对象集合
        List<XWPFTable> tables = document.getTables();
        // 循环word中的所有表格
        for (int k = 0; k < tables.size(); k++) {
            // 获取单个表格
            XWPFTable table = tables.get(k);
            // 获取要替换的数据
            TableForm tableForm = list.get(k);
            Integer headerIndex = tableForm.getStartLine();
            List<String[]> tableList = tableForm.getData();
            if (null == tableList) {
                return;
            }
            XWPFTableRow copyRow = table.getRow(headerIndex);
            List<XWPFTableCell> cellList = copyRow.getTableCells();
            if (null == cellList) {
                break;
            }
            //遍历要添加的数据的list
            for (int i = 0; i < tableList.size(); i++) {
                //插入一行
                XWPFTableRow targetRow = table.insertNewTableRow(headerIndex + 1 + i);
                //复制行属性
                targetRow.getCtRow().setTrPr(copyRow.getCtRow().getTrPr());

                String[] strings = tableList.get(i);
                for (int j = 0; j < strings.length; j++) {
                    XWPFTableCell sourceCell = cellList.get(j);
                    //插入一个单元格
                    XWPFTableCell targetCell = targetRow.addNewTableCell();
                    //复制列属性
                    targetCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
                    targetCell.setText(strings[j]);
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
            String key = "${" + textSet.getKey() + "}";
            if (value.indexOf(key) != -1) {
                value = textSet.getValue();
            }
        }
        //模板未匹配到区域替换为空
        if (checkText(value)) {
            value = "";
        }
        return value;
    }

    /**
     * 遍历表格,并替换模板
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

    static CellReference setTitleInDataSheet(XWPFChart chart, String title, int column) throws Exception {
        XSSFWorkbook workbook = chart.getWorkbook();
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFRow row = sheet.getRow(0);
        if (row == null)
            row = sheet.createRow(0);
        XSSFCell cell = row.getCell(column);
        if (cell == null)
            cell = row.createCell(column);
        cell.setCellValue(title);
        return new CellReference(sheet.getSheetName(), 0, column, true, true);
    }
}