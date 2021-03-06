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
package com.github.microprograms.poi_template.policy;

import org.apache.poi.xwpf.usermodel.XWPFRun;

import com.github.microprograms.poi_template.XWPFTemplate;
import com.github.microprograms.poi_template.template.ElementTemplate;
import com.github.microprograms.poi_template.template.run.RunTemplate;
import com.github.microprograms.poi_template.xwpf.BodyContainer;
import com.github.microprograms.poi_template.xwpf.BodyContainerFactory;

public class CommentPolicy implements RenderPolicy {

    @Override
    public void render(ElementTemplate eleTemplate, Object data, XWPFTemplate template) {
        RunTemplate runTemplate = (RunTemplate) eleTemplate;
        XWPFRun run = runTemplate.getRun();
        BodyContainer bodyContainer = BodyContainerFactory.getBodyContainer(run);
        bodyContainer.clearPlaceholder(run);
    }

}
