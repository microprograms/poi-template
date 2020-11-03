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

import com.github.microprograms.poi_template.data.style.Style;

/**
 * hyper link text
 */
public class HyperlinkTextRenderData extends TextRenderData {

    private static final long serialVersionUID = 1L;

    /**
     * link format: https://github.com/microprograms <br/>
     * 
     * mail format: mailto:xuzewei_2012@126.com?subject=poi-template <br/>
     * 
     * anchor format：anchor:AnchorName
     */
    private String url;

    HyperlinkTextRenderData() {
    }

    public HyperlinkTextRenderData(String text, String url) {
        super(text);
        this.url = url;
        this.style = Style.builder().buildColor("0000FF").buildUnderLine().build();
    }

    public String getUrl() {
        return url;
    }

    public void setUrl(String url) {
        this.url = url;
    }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append("[").append(text).append("](").append(url).append(")");
        return sb.toString();
    }

}
