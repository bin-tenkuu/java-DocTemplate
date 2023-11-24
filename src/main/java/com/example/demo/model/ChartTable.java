package com.example.demo.model;

import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import lombok.val;

import java.util.HashMap;

@Getter
@Setter
@NoArgsConstructor
public class ChartTable {
    private String title;
    private ChartColumn<String> xAxis = new ChartColumn<>("x轴");
    private HashMap<String, ChartColumn<Number>> yAxis = new HashMap<>();

    public ChartTable(String title) {
        this.title = title;
    }

    public ChartTable setTitle(String title) {
        this.title = title;
        return this;
    }

    public ChartColumn<Number> newYAxis(String title) {
        val column = new ChartColumn<Number>();
        yAxis.put(title, column);
        return column;
    }

    public ChartColumn<Number> getYAxis(String title) {
        return yAxis.get(title);
    }

}
