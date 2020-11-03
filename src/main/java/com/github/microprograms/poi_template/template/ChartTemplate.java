/*
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.github.microprograms.poi_template.template;

import java.util.List;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.XDDFArea3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFAreaChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFBar3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFLine3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFPie3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFPieChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFRadarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFScatterChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFSurface3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFSurfaceChartData;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import com.github.microprograms.poi_template.config.Configure;
import com.github.microprograms.poi_template.policy.RenderPolicy;
import com.github.microprograms.poi_template.render.processor.Visitor;

/**
 * chart docx template element: XWPFChart
 */
public class ChartTemplate extends ElementTemplate {

    protected XWPFChart chart;
    protected ChartTypes chartType;
    protected XWPFRun run;

    public ChartTemplate(String tagName, XWPFChart chart, XWPFRun run) {
        this.tagName = tagName;
        this.chart = chart;
        this.run = run;
        this.chartType = readChartType(this.chart);
    }

    private ChartTypes readChartType(XWPFChart chart) {
        List<XDDFChartData> chartSeries = chart.getChartSeries();
        if (CollectionUtils.isEmpty(chartSeries)) return null;
        XDDFChartData chartData = chartSeries.get(0);
        ChartTypes chartType = null;
        if (chartData.getClass() == XDDFAreaChartData.class) {
            chartType = ChartTypes.AREA;
        } else if (chartData.getClass() == XDDFArea3DChartData.class) {
            chartType = ChartTypes.AREA3D;
        } else if (chartData.getClass() == XDDFBarChartData.class) {
            chartType = ChartTypes.BAR;
        } else if (chartData.getClass() == XDDFBar3DChartData.class) {
            chartType = ChartTypes.BAR3D;
        } else if (chartData.getClass() == XDDFLineChartData.class) {
            chartType = ChartTypes.LINE;
        } else if (chartData.getClass() == XDDFLine3DChartData.class) {
            chartType = ChartTypes.LINE3D;
        } else if (chartData.getClass() == XDDFPieChartData.class) {
            chartType = ChartTypes.PIE;
        } else if (chartData.getClass() == XDDFPie3DChartData.class) {
            chartType = ChartTypes.PIE3D;
        } else if (chartData.getClass() == XDDFRadarChartData.class) {
            chartType = ChartTypes.RADAR;
        } else if (chartData.getClass() == XDDFScatterChartData.class) {
            chartType = ChartTypes.SCATTER;
        } else if (chartData.getClass() == XDDFSurfaceChartData.class) {
            chartType = ChartTypes.SURFACE;
        } else if (chartData.getClass() == XDDFSurface3DChartData.class) {
            chartType = ChartTypes.SURFACE3D;
        }
        return chartType;
    }

    public XWPFChart getChart() {
        return chart;
    }

    public void setChart(XWPFChart chart) {
        this.chart = chart;
    }

    public XWPFRun getRun() {
        return run;
    }

    public void setRun(XWPFRun run) {
        this.run = run;
    }

    @Override
    public void accept(Visitor visitor) {
        visitor.visit(this);
    }

    public RenderPolicy findPolicy(Configure config) {
        RenderPolicy policy = config.getCustomPolicy(tagName);
        if (null == policy) policy = config.getChartPolicy(chartType);
        return null == policy ? config.getTemplatePolicy(this.getClass()) : policy;
    }

    @Override
    public String toString() {
        return "Chart::" + source;
    }

}
