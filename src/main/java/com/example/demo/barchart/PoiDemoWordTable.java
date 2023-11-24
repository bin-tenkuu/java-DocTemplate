package com.example.demo.barchart;

import com.example.demo.model.ChartTable;
import com.example.demo.model.WordParam;
import com.example.demo.model.WordParams;
import com.example.demo.util.PoiWordTools;
import lombok.val;

import java.io.File;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

public class PoiDemoWordTable {

    private static WordParams buildMap() throws IOException {
        val params = WordParams.create();

        params.setText("名字", "系统");
        params.setParam("图片", WordParam.image(new File("./config/img.png")));

        ChartTable chartTable = params.addChart("params")
                .setTitle("更新标题-占基金资产净值比例(%)");
        chartTable.getXAxis()
                .addAllData("材料费用", "出差费用", "住宿费用");
        chartTable.newYAxis("金额2")
                .addAllData(5002, 3002, 3002);
        chartTable.newYAxis("金额")
                .addAllData(500, 300, 300);

        chartTable = params.addChart("系列 1")
                .setTitle("更新标题-系列 1");
        chartTable.getXAxis()
                .addAllData("材料费用", "出差费用", "住宿费用");
        chartTable.newYAxis("金额2")
                .addAllData(5002, 3002, 3002);
        chartTable.newYAxis("金额")
                .addAllData(500, 300, 300);

        chartTable = params.addChart("图表标题")
                .setTitle("更新标题-我是修改后的标题");
        chartTable.getXAxis()
                .addAllData("老张2", "老李3", "老刘4");
        chartTable.newYAxis("存款")
                .setDateList(random(3));
        chartTable.newYAxis("存款2")
                .setDateList(random(3));

        chartTable = params.addChart("标题4")
                .setTitle("更新标题-标题4");
        chartTable.newYAxis("占基金资产净值比例22222（%）")
                .addAllData(random(3));
        chartTable.newYAxis("1")
                .setTitle("比例1")
                .setDateList(random(3));
        chartTable.newYAxis("2")
                .setTitle("第二比例")
                .setDateList(random(3));

        chartTable = params.addChart("销售额")
                .setTitle("更新标题-销售额");
        chartTable.getXAxis()
                .addAllData("材料费用", "出差费用", "住宿费用");
        chartTable.newYAxis("金额2")
                .addAllData(5002, 3002, 3002);

        chartTable = params.addChart("标题6")
                .setTitle("更新标题-标题6");
        chartTable.getXAxis()
                .addAllData("通辽", "呼和浩特", "锡林郭勒",
                        "阿拉善", "巴彦淖尔", "兴安",
                        "乌兰察布", "乌海", "赤峰",
                        "包头", "呼伦贝尔", "鄂尔多斯");
        chartTable.newYAxis("投诉受理量（次）")
                .setDateList(random(12));
        chartTable.newYAxis("预处理拦截工单量（次）")
                .setDateList(random(12));
        chartTable.newYAxis("拦截率")
                .setDateList(random(12));

        return params;
    }

    /**
     * 模拟系统中获取的数据列表
     */
    private static List<Number> random(int n) {
        val numbers = new Number[n];
        for (int i = 0; i < n; i++) {
            numbers[i] = (int) (1 + Math.random() * 100);
        }
        return Arrays.asList(numbers);
    }

    public static void main(String[] args) throws Exception {
        val map = buildMap();
        PoiWordTools.buildDoc(
                map,
                new File("./config/test.docx"),
                new File("./test.docx")
        );
    }

}





















