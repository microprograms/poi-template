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

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;

import com.github.microprograms.poi_template.data.style.ParagraphStyle;
import com.github.microprograms.poi_template.data.style.Style;

/**
 * Factory to create {@link ParagraphRenderData}
 */
public class Paragraphs {

    private Paragraphs() {
    }

    public static ParagraphBuilder of() {
        ParagraphBuilder inst = new ParagraphBuilder();
        return inst;
    }

    public static ParagraphBuilder of(String text) {
        return Paragraphs.of().addText(text);
    }

    public static ParagraphBuilder of(TextRenderData text) {
        return Paragraphs.of().addText(text);
    }

    public static ParagraphBuilder of(PictureRenderData picture) {
        return Paragraphs.of().addPicture(picture);
    }

    /**
     * Builder to build {@link ParagraphRenderData}
     */
    public static class ParagraphBuilder implements RenderDataBuilder<ParagraphRenderData> {

        private List<RenderData> contents = new ArrayList<>();
        private ParagraphStyle paragraphStyle;

        private ParagraphBuilder() {
        }

        public ParagraphBuilder addText(TextRenderData text) {
            contents.add(text);
            return this;
        }

        public ParagraphBuilder addText(String text) {
            contents.add(Texts.of(text).create());
            return this;
        }

        public ParagraphBuilder addPicture(PictureRenderData picture) {
            contents.add(picture);
            return this;
        }

        public ParagraphBuilder paraStyle(ParagraphStyle style) {
            this.paragraphStyle = style;
            return this;
        }

        public ParagraphBuilder glyphStyle(Style style) {
            if (null == this.paragraphStyle) {
                this.paragraphStyle = ParagraphStyle.builder().withGlyphStyle(style).build();
            } else {
                this.paragraphStyle.setGlyphStyle(style);
            }
            return this;
        }

        public ParagraphBuilder left() {
            if (null == this.paragraphStyle) {
                this.paragraphStyle = ParagraphStyle.builder().withAlign(ParagraphAlignment.LEFT).build();
            } else {
                this.paragraphStyle.setAlign(ParagraphAlignment.LEFT);
            }
            return this;
        }

        public ParagraphBuilder center() {
            if (null == this.paragraphStyle) {
                this.paragraphStyle = ParagraphStyle.builder().withAlign(ParagraphAlignment.CENTER).build();
            } else {
                this.paragraphStyle.setAlign(ParagraphAlignment.CENTER);
            }
            return this;
        }

        public ParagraphBuilder right() {
            if (null == this.paragraphStyle) {
                this.paragraphStyle = ParagraphStyle.builder().withAlign(ParagraphAlignment.RIGHT).build();
            } else {
                this.paragraphStyle.setAlign(ParagraphAlignment.RIGHT);
            }
            return this;
        }

        @Override
        public ParagraphRenderData create() {
            ParagraphRenderData data = new ParagraphRenderData();
            data.setContents(contents);
            data.setParagraphStyle(paragraphStyle);
            return data;
        }
    }

}
