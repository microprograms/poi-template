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

import com.github.microprograms.poi_template.data.builder.StyleBuilder;

/**
 * 超链接
 */
public class HyperLinkTextRenderData extends TextRenderData {

    /**
     * 超链接: https://github.com/microprograms <br/>
     * 发送邮件链接: mailto:xuzewei_2012@126.com?subject=poi-template <br/>
     * 锚点: anchor:AnchorName
     */
    private String url;

    public HyperLinkTextRenderData(String text, String url) {
        super(text);
        this.url = url;
        // 链接默认蓝色加下划线
        this.style = StyleBuilder.newBuilder().buildColor("0000FF").buildUnderLine().build();
    }

    public String getUrl() {
        return url;
    }

    public void setUrl(String url) {
        this.url = url;
    }

}
