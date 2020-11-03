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

import java.util.List;
import java.util.Objects;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import com.github.microprograms.poi_template.data.ParagraphRenderData;
import com.github.microprograms.poi_template.data.PictureRenderData;
import com.github.microprograms.poi_template.data.RenderData;
import com.github.microprograms.poi_template.data.TextRenderData;
import com.github.microprograms.poi_template.data.style.ParagraphStyle;
import com.github.microprograms.poi_template.render.RenderContext;
import com.github.microprograms.poi_template.util.ParagraphUtils;
import com.github.microprograms.poi_template.util.StyleUtils;
import com.github.microprograms.poi_template.xwpf.XWPFParagraphWrapper;

/**
 * paragraph render
 */
public class ParagraphRenderPolicy extends AbstractRenderPolicy<ParagraphRenderData> {

    @Override
    protected boolean validate(ParagraphRenderData data) {
        return null != data && !data.getContents().isEmpty();
    }

    @Override
    protected void afterRender(RenderContext<ParagraphRenderData> context) {
        clearPlaceholder(context, false);
    }

    @Override
    public void doRender(RenderContext<ParagraphRenderData> context) throws Exception {
        Helper.renderParagraph(context.getRun(), context.getData());

    }

    public static class Helper {

        public static void renderParagraph(XWPFRun run, ParagraphRenderData data) throws Exception {
            renderParagraph(run, data, null);
        }

        public static void renderParagraph(XWPFRun run, ParagraphRenderData data,
                List<ParagraphStyle> defaultControlStyles) throws Exception {
            List<RenderData> contents = data.getContents();
            XWPFParagraph paragraph = (XWPFParagraph) run.getParent();
            styleParagraphWithDefaultStyle(paragraph, defaultControlStyles);
            StyleUtils.styleParagraph(paragraph, data.getParagraphStyle());

            XWPFParagraphWrapper parentContext = new XWPFParagraphWrapper(paragraph);
            for (RenderData content : contents) {
                XWPFRun fragment = parentContext.insertNewRun(ParagraphUtils.getRunPos(run));
                StyleUtils.styleRun(fragment, run);
                if (content instanceof TextRenderData) {
                    styleRunWithDefaultStyle(fragment, defaultControlStyles);
                    StyleUtils.styleRun(fragment,
                            null == data.getParagraphStyle() ? null : data.getParagraphStyle().getDefaultTextStyle());
                    TextRenderPolicy.Helper.renderTextRun(fragment, content);
                } else if (content instanceof PictureRenderData) {
                    PictureRenderPolicy.Helper.renderPicture(fragment, (PictureRenderData) content);
                }
            }

        }

        private static void styleRunWithDefaultStyle(XWPFRun fragment, List<ParagraphStyle> defaultControlStyles) {
            if (null != defaultControlStyles) {
                defaultControlStyles.stream().filter(Objects::nonNull).forEach(style -> {
                    StyleUtils.styleRun(fragment, style.getDefaultTextStyle());
                });
            }
        }

        private static void styleParagraphWithDefaultStyle(XWPFParagraph paragraph,
                List<ParagraphStyle> defaultControlStyles) {
            if (null != defaultControlStyles) {
                defaultControlStyles.stream().filter(Objects::nonNull)
                        .forEach(style -> StyleUtils.styleParagraph(paragraph, style));
            }
        }
    }

}
