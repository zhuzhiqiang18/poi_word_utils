package com.lh.word.form;


import lombok.Data;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.XDDFLineProperties;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.MarkerStyle;

import java.util.ArrayList;
import java.util.List;

/**
 * Copyright (C), 2006-2010, ChengDu ybya info. Co., Ltd.
 * FileName: ScatterChartForm.java
 *
 * @author lh
 * @version 1.0.0
 * @Date 2021/02/03 11:13
 */
@Data
public class ScatterChartForm extends ChartFrom {

    private List<AreaData> lists = new ArrayList<>();

    // 标记大小 默认为6
    private Short markerSize = 6;

    // 标记样式 默认圆
    private MarkerStyle style = MarkerStyle.CIRCLE;

    // 是否弯曲 默认不
    private Boolean smooth;

    // 是否自动生成颜色
    private Boolean varyColors;

    private Boolean isShowXGrid = true;//是否显示X轴网格线
    private Boolean isShowYGrid = true;//是否显示Y轴网格线

    private Boolean isShowLegend = false;

    private double xMajorUnit = 3;//x坐标步长

    private double yMajorUnit = 20;//y坐标步长



    public XDDFLineProperties line(){
        XDDFLineProperties line = new XDDFLineProperties();
        line.setFillProperties(new XDDFSolidFillProperties(XDDFColor.from(new byte[] {(byte)228,(byte)231,(byte)237})));
        line.setWidth(0.5);
        return line;
    }

    @Data
    public static class AreaData {
        // X轴数据
        private Integer[] bottomData;

        // Y轴数据
        private Integer[] leftData;

        // 点的颜色,可以为空 创建方式为 XDDFColor.from(new byte[]{(byte)0xFF, (byte)0xE1, (byte)0xFF})
        private XDDFColor xddfColor;

        // 系列名称
        private String title;

    }


}