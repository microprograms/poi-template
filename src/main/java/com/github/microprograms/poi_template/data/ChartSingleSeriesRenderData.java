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
package com.github.microprograms.poi_template.data;

/**
 * 单系列图表：环形图、饼图
 */
public class ChartSingleSeriesRenderData implements RenderData {

    private String chartTitle;
    private String[] categories;
    private SeriesRenderData seriesData;

    public String getChartTitle() {
        return chartTitle;
    }

    public void setChartTitle(String chartTitle) {
        this.chartTitle = chartTitle;
    }

    public String[] getCategories() {
        return categories;
    }

    public void setCategories(String[] categories) {
        this.categories = categories;
    }

    public SeriesRenderData getSeriesData() {
        return seriesData;
    }

    public void setSeriesData(SeriesRenderData seriesData) {
        this.seriesData = seriesData;
    }

}
