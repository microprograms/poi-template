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

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.InputStream;

import com.github.microprograms.poi_template.util.BufferedImageUtils;
import com.github.microprograms.poi_template.util.ByteUtils;

/**
 * Factory method to build {@link PictureRenderData} instances.
 */
public class Pictures {
    private Pictures() {
    }

    public static PictureBuilder ofLocal(String src) {
        return Pictures.ofBytes(ByteUtils.getLocalByteArray(new File(src)), PictureType.suggestFileType(src));
    }

    public static PictureBuilder ofUrl(String url, PictureType pictureType) {
        return ofBytes(ByteUtils.getUrlByteArray(url), pictureType);
    }

    public static PictureBuilder ofStream(InputStream inputStream, PictureType pictureType) {
        return ofBytes(ByteUtils.toByteArray(inputStream), pictureType);
    }

    public static PictureBuilder ofBufferedImage(BufferedImage image, PictureType pictureType) {
        return ofBytes(BufferedImageUtils.getBufferByteArray(image, pictureType.format()), pictureType);
    }

    public static PictureBuilder ofBytes(byte[] bytes, PictureType pictureType) {
        PictureBuilder inst = new PictureBuilder(pictureType, bytes);
        return inst;
    }

    /**
     * Builder to build {@link PictureRenderData} instances.
     *
     */
    public static class PictureBuilder implements RenderDataBuilder<PictureRenderData> {

        private PictureRenderData data;

        private PictureBuilder(PictureType pictureType, byte[] bytes) {
            data = new PictureRenderData(0, 0, pictureType, bytes);
        }

        public PictureBuilder size(int width, int height) {
            data.setWidth(width);
            data.setHeight(height);
            return this;
        }

        public PictureBuilder altMeta(String altMeta) {
            data.setAltMeta(altMeta);
            return this;
        }

        @Override
        public PictureRenderData create() {
            return data;
        }
    }

}
