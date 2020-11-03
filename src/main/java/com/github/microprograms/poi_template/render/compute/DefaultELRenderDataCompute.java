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

import com.github.microprograms.poi_template.exception.ExpressionEvalException;
import com.github.microprograms.poi_template.expression.DefaultEL;

/**
 * default expression compute
 */
public class DefaultELRenderDataCompute implements RenderDataCompute {

    private DefaultEL elObject;
    private boolean isStrict;

    public DefaultELRenderDataCompute(Object root, boolean isStrict) {
        elObject = DefaultEL.create(root);
        this.isStrict = isStrict;
    }

    @Override
    public Object compute(String el) {
        try {
            return elObject.eval(el);
        } catch (ExpressionEvalException e) {
            if (isStrict) throw e;
            // Cannot calculate the expression, the default returns null
            return null;
        }
    }

}
