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

package com.github.microprograms.poi_template.render.processor;

import org.apache.commons.lang3.ClassUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.microprograms.poi_template.XWPFTemplate;
import com.github.microprograms.poi_template.exception.RenderException;
import com.github.microprograms.poi_template.policy.DocxRenderPolicy;
import com.github.microprograms.poi_template.policy.RenderPolicy;
import com.github.microprograms.poi_template.render.compute.RenderDataCompute;
import com.github.microprograms.poi_template.resolver.Resolver;
import com.github.microprograms.poi_template.template.ChartTemplate;
import com.github.microprograms.poi_template.template.ElementTemplate;
import com.github.microprograms.poi_template.template.PictureTemplate;
import com.github.microprograms.poi_template.template.run.RunTemplate;

public class ElementProcessor extends DefaultTemplateProcessor {

    private static Logger logger = LoggerFactory.getLogger(ElementProcessor.class);

    public ElementProcessor(XWPFTemplate template, Resolver resolver, RenderDataCompute renderDataCompute) {
        super(template, resolver, renderDataCompute);
    }

    @Override
    public void visit(PictureTemplate pictureTemplate) {
        visit((ElementTemplate) pictureTemplate);
    }

    @Override
    public void visit(ChartTemplate chartTemplate) {
        visit((ElementTemplate) chartTemplate);
    }

    @Override
    public void visit(RunTemplate runTemplate) {
        visit((ElementTemplate) runTemplate);
    }

    void visit(ElementTemplate eleTemplate) {
        RenderPolicy policy = eleTemplate.findPolicy(template.getConfig());
        if (null == policy) {
            throw new RenderException("Cannot find render policy: [" + eleTemplate.getTagName() + "]");
        }
        if (policy instanceof DocxRenderPolicy) return;
        logger.info("Start render Template {}, Sign:{}, policy:{}", eleTemplate, eleTemplate.getSign(),
                ClassUtils.getShortClassName(policy.getClass()));
        policy.render(eleTemplate, renderDataCompute.compute(eleTemplate.getTagName()), template);
    }

}
