package com.example.demo.util;

import com.example.demo.model.ChartColumn;
import com.example.demo.model.ChartTable;
import com.example.demo.model.WordParam;
import com.example.demo.model.WordParams;
import lombok.extern.slf4j.Slf4j;
import lombok.val;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * poi生成word的工具类
 * 针对于模板中的图表是静态的，也就是模板中的图表长什么样子不会根据数据而改变
 */
@Slf4j
public class PoiWordTools {

    private enum ReplaceType {
        NotFound,
        $,
        Start,
    }

    private static void replaceParagraph(XWPFDocument doc, WordParams params)
            throws IOException, InvalidFormatException {
        for (XWPFParagraph paragraph : doc.getParagraphs()) {
            val runs = paragraph.getRuns();

            val prevText = new StringBuilder();
            XWPFRun replaceRun = null;
            int replaceIndex = 0;
            ReplaceType replaceType = ReplaceType.NotFound;
            val key = new StringBuilder();

            for (int index = 0; index < runs.size(); index++) {
                val run = runs.get(index);
                // run.getCTR().sizeOfTArray();
                val text = run.getText(0);
                if (text == null || text.isEmpty()) {
                    continue;
                }
                switch (replaceType) {
                    case NotFound:
                        replaceIndex = text.indexOf('$');
                        if (replaceIndex < 0) {
                            continue;
                        }
                        replaceRun = run;
                        prevText.append(text, 0, replaceIndex);
                        replaceIndex++;
                        replaceType = ReplaceType.$;
                    case $:
                        if (replaceIndex == text.length()) {
                            replaceIndex = 0;
                            continue;
                        }
                        if (text.charAt(replaceIndex) != '{') {
                            replaceRun = null;
                            prevText.setLength(0);
                            replaceType = ReplaceType.NotFound;
                            continue;
                        }
                        replaceIndex++;
                        replaceType = ReplaceType.Start;
                    case Start:
                        if (replaceIndex == text.length()) {
                            if (replaceRun != run) {
                                paragraph.removeRun(index);
                                index--;
                            }
                            replaceIndex = 0;
                            continue;
                        }
                        var replaceEnd = text.indexOf('}', replaceIndex);
                        if (replaceEnd < 0) {
                            key.append(text, replaceIndex, text.length());
                            if (replaceRun != run) {
                                paragraph.removeRun(index);
                                index--;
                            }
                            continue;
                        }
                        key.append(text, replaceIndex, replaceEnd);
                        val value = params.getParam(key.toString());
                        if (value == null) {
                            prevText.append("${");
                            prevText.append(key);
                            prevText.append("}");
                            replaceRun.setText(prevText.toString(), 0);
                        } else if (value instanceof WordParam.Text msg) {
                            prevText.append(msg.getMsg());
                            replaceRun.setText(prevText.toString(), 0);
                        } else if (value instanceof WordParam.Image image) {
                            replaceRun.setText(prevText.toString(), 0);
                            key.append(".jpg");
                            replaceRun.addPicture(
                                    image.getInputStream(),
                                    Document.PICTURE_TYPE_JPEG,
                                    key.toString(),
                                    image.getWidth(),
                                    image.getHeight()
                            );
                        }
                        prevText.setLength(0);
                        key.setLength(0);
                        replaceEnd++;
                        if (replaceEnd < text.length()) {
                            run.setText(text.substring(replaceEnd), 0);
                            index--;
                        } else if (replaceRun != run) {
                            paragraph.removeRun(index);
                            index--;
                        }
                        replaceRun = null;
                        replaceType = ReplaceType.NotFound;
                }
            }
        }
    }

    private static void replaceChart(XWPFDocument doc, WordParams params) {
        for (POIXMLDocumentPart poixmlDocumentPart : doc.getRelations()) {
            if (poixmlDocumentPart instanceof XWPFChart chart) {
                String key = PoiWordTools.getBarTitle(chart);
                val chartTable = params.getChart(key);
                if (chartTable != null) {
                    replaceCharts(chart, chartTable);
                }
            }
        }

    }

    public static boolean buildDoc(WordParams params, File from, File to) {
        try (val in = new FileInputStream(from)) {
            try (val doc = new XWPFDocument(in)) {
                replaceParagraph(doc, params);
                replaceChart(doc, params);
                try (val out = new FileOutputStream(to)) {
                    doc.write(out);
                }
            }
            return true;
        } catch (IOException | InvalidFormatException e) {
            log.error("生成word失败", e);
            return false;
        }
    }

    public static String getBarTitle(XWPFChart chart) {
        XDDFTitle title = chart.getTitle();
        if (title != null) {
            return title.getBody().getParagraph(0).getText();
        }
        return null;
    }

    /**
     * 调用替换柱状图、折线图组合数据
     */
    public static void replaceCharts(XWPFChart chart, ChartTable chartTable) {
        // 设置标题
        chart.setTitleText(chartTable.getTitle());

        val x = fromString(chartTable.getXAxis());

        for (XDDFChartData chartData : chart.getChartSeries()) {
            for (int i = 0; i < chartData.getSeriesCount(); i++) {
                val series = chartData.getSeries(i);
                val title = PoiUtil.getTitle(series);
                if (title == null) {
                    continue;
                }
                val yAxis = chartTable.getYAxis(title);
                if (yAxis == null) {
                    continue;
                }
                if (!title.equals(yAxis.getTitle())) {
                    series.setTitle(yAxis.getTitle(), null);
                }
                val dataSource = fromNumber(yAxis);
                series.replaceData(x, dataSource);
                series.plot();
            }
        }
    }

    public static XDDFCategoryDataSource fromString(ChartColumn<String> column) {
        return XDDFDataSourcesFactory.fromArray(column.toArray(new String[0]));
    }

    public static XDDFNumericalDataSource<Number> fromNumber(ChartColumn<Number> column) {
        return XDDFDataSourcesFactory.fromArray(column.toArray(new Number[0]));
    }

}















