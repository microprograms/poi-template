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

import java.util.HashSet;
import java.util.List;
import java.util.Set;

import com.github.microprograms.poi_template.XWPFTemplate;
import com.github.microprograms.poi_template.render.compute.RenderDataCompute;
import com.github.microprograms.poi_template.resolver.Resolver;
import com.github.microprograms.poi_template.template.ChartTemplate;
import com.github.microprograms.poi_template.template.InlineIterableTemplate;
import com.github.microprograms.poi_template.template.IterableTemplate;
import com.github.microprograms.poi_template.template.MetaTemplate;
import com.github.microprograms.poi_template.template.PictureTemplate;
import com.github.microprograms.poi_template.template.run.RunTemplate;
import com.github.microprograms.poi_template.xwpf.XWPFTextboxContent;

public abstract class DefaultTemplateProcessor implements Visitor {

    protected XWPFTemplate template;
    protected final RenderDataCompute renderDataCompute;
    protected final Resolver resolver;;

    public DefaultTemplateProcessor(XWPFTemplate template, final Resolver resolver,
            final RenderDataCompute renderDataCompute) {
        this.template = template;
        this.resolver = resolver;
        this.renderDataCompute = renderDataCompute;
    }

    protected void visitOther(MetaTemplate template) {
        // no-op
    }

    @SuppressWarnings("deprecation")
    public void process(List<MetaTemplate> templates) {
        // process in order( or sort first)
        templates.forEach(template -> template.accept(this));
        Set<XWPFTextboxContent> textboxs = new HashSet<>();
        templates.forEach(template -> {
            if (template instanceof RunTemplate) {
                if (((RunTemplate) template).getRun().getParagraph().getBody() instanceof XWPFTextboxContent) {
                    textboxs.add((XWPFTextboxContent) ((RunTemplate) template).getRun().getParagraph().getBody());
                }
            }
        });
        textboxs.forEach(content -> {
            content.getXmlObject().set(content.getCTTxbxContent());
        });
    }

    @Override
    public void visit(InlineIterableTemplate iterableTemplate) {
        visitOther(iterableTemplate);
    }

    @Override
    public void visit(RunTemplate runTemplate) {
        visitOther(runTemplate);
    }

    @Override
    public void visit(PictureTemplate pictureTemplate) {
        visitOther(pictureTemplate);
    }

    @Override
    public void visit(ChartTemplate chartTemplate) {
        visitOther(chartTemplate);
    }

    @Override
    public void visit(IterableTemplate iterableTemplate) {
        visitOther(iterableTemplate);
    }

}
