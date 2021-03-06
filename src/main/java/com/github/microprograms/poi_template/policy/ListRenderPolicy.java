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

import org.apache.poi.xwpf.usermodel.XWPFRun;

import com.github.microprograms.poi_template.data.MiniTableRenderData;
import com.github.microprograms.poi_template.data.NumbericRenderData;
import com.github.microprograms.poi_template.data.PictureRenderData;
import com.github.microprograms.poi_template.data.TextRenderData;
import com.github.microprograms.poi_template.render.RenderContext;
import com.github.microprograms.poi_template.util.StyleUtils;
import com.github.microprograms.poi_template.xwpf.BodyContainer;
import com.github.microprograms.poi_template.xwpf.BodyContainerFactory;

public class ListRenderPolicy extends AbstractRenderPolicy<List<Object>> {

    @Override
    protected boolean validate(List<Object> data) {
        return (null != data && !data.isEmpty());
    }

    @Override
    public void doRender(RenderContext<List<Object>> context) throws Exception {
        XWPFRun run = context.getRun();
        BodyContainer bodyContainer = BodyContainerFactory.getBodyContainer(run);
        List<Object> datas = context.getData();
        for (Object data : datas) {
            if (data instanceof TextRenderData) {
                XWPFRun createRun = bodyContainer.insertNewParagraph(run).createRun();
                StyleUtils.styleRun(createRun, run);
                TextRenderPolicy.Helper.renderTextRun(createRun, data);
            } else if (data instanceof MiniTableRenderData) {
                MiniTableRenderPolicy.Helper.renderMiniTable(run, (MiniTableRenderData) data);
            } else if (data instanceof NumbericRenderData) {
                NumbericRenderPolicy.Helper.renderNumberic(run, (NumbericRenderData) data);
            } else if (data instanceof PictureRenderData) {
                PictureRenderPolicy.Helper.renderPicture(bodyContainer.insertNewParagraph(run).createRun(),
                        (PictureRenderData) data);
            }
        }
    }

    @Override
    protected void afterRender(RenderContext<List<Object>> context) {
        clearPlaceholder(context, true);
    }

}
