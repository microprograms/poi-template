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
package com.github.microprograms.poi_template.render.compute;

public class DefaultRenderDataComputeFactory implements RenderDataComputeFactory {

    private final boolean isStrict;

    public DefaultRenderDataComputeFactory(boolean isStrict) {
        this.isStrict = isStrict;
    }

    @Override
    public RenderDataCompute newCompute(Object model) {
        return new DefaultELRenderDataCompute(model, isStrict);
    }

}
