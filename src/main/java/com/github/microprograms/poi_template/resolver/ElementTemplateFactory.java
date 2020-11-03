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
package com.github.microprograms.poi_template.resolver;

import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import com.github.microprograms.poi_template.template.ChartTemplate;
import com.github.microprograms.poi_template.template.PictureTemplate;
import com.github.microprograms.poi_template.template.run.RunTemplate;

public interface ElementTemplateFactory {

    RunTemplate createRunTemplate(String tag, XWPFRun run);

    PictureTemplate createPicureTemplate(String tag, XWPFPicture pic);

    ChartTemplate createChartTemplate(String tag, XWPFChart chart, XWPFRun run);

}
