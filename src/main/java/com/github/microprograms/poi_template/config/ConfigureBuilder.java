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
package com.github.microprograms.poi_template.config;

import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.Map;

import com.github.microprograms.poi_template.config.Configure.CustomPolicyFinder;
import com.github.microprograms.poi_template.config.Configure.ELMode;
import com.github.microprograms.poi_template.config.Configure.ValidErrorHandler;
import com.github.microprograms.poi_template.policy.RenderPolicy;
import com.github.microprograms.poi_template.render.compute.ELObjectRenderDataCompute;
import com.github.microprograms.poi_template.render.compute.RenderDataComputeFactory;
import com.github.microprograms.poi_template.render.compute.SpELRenderDataCompute;
import com.github.microprograms.poi_template.resolver.ElementTemplateFactory;
import com.github.microprograms.poi_template.template.MetaTemplate;
import com.github.microprograms.poi_template.util.RegexUtils;

import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;

public class ConfigureBuilder {
    private Configure config = new Configure();

    public ConfigureBuilder gramer(String prefix, String suffix) {
        config.gramerPrefix = prefix;
        config.gramerSuffix = suffix;
        return this;
    }

    public ConfigureBuilder iterableLeft(char c) {
        config.iterable = Pair.of(c, config.iterable.getRight());
        return this;
    }

    public ConfigureBuilder grammerRegex(String reg) {
        config.grammerRegex = reg;
        return this;
    }

    public ConfigureBuilder useSpringEL() {
        return useSpringEL(root -> new SpELRenderDataCompute(root, new HashMap<>()));
    }

    public ConfigureBuilder useSpringEL(Map<String, Method> spELFunction) {
        return useSpringEL(root -> new SpELRenderDataCompute(root, spELFunction));
    }

    public ConfigureBuilder useSpringEL(RenderDataComputeFactory renderDataComputeFactory) {
        config.elMode = ELMode.SPEL_MODE;
        return _setRenderDataComputeFactory(renderDataComputeFactory);
    }

    public ConfigureBuilder useDefaultEL(boolean isStrict) {
        config.elMode = isStrict ? ELMode.POI_TL_STICT_MODE : ELMode.POI_TL_STANDARD_MODE;
        return _setRenderDataComputeFactory(root -> new ELObjectRenderDataCompute(root, isStrict));
    }

    private ConfigureBuilder _setRenderDataComputeFactory(RenderDataComputeFactory renderDataComputeFactory) {
        config.renderDataComputeFactory = renderDataComputeFactory;
        return this;
    }

    public ConfigureBuilder setValidErrorHandler(ValidErrorHandler handler) {
        config.handler = handler;
        return this;
    }

    public ConfigureBuilder setElementTemplateFactory(ElementTemplateFactory elementTemplateFactory) {
        config.elementTemplateFactory = elementTemplateFactory;
        return this;
    }

    public ConfigureBuilder addPlugin(char c, RenderPolicy policy) {
        config.plugin(c, policy);
        return this;
    }

    public ConfigureBuilder addPlugin(Class<? extends MetaTemplate> clazz, RenderPolicy policy) {
        config.plugin(clazz, policy);
        return this;
    }

    public ConfigureBuilder addPlugin(ChartTypes chartType, RenderPolicy policy) {
        config.plugin(chartType, policy);
        return this;
    }

    public ConfigureBuilder setCustomPolicyFinder(CustomPolicyFinder customPolicyFinder) {
        config.setCustomPolicyFinder(customPolicyFinder);
        return this;
    }

    public ConfigureBuilder bind(String tagName, RenderPolicy policy) {
        config.customPolicy(tagName, policy);
        return this;
    }

    public Configure build() {
        if (config.elMode == ELMode.SPEL_MODE) {
            config.grammerRegex = RegexUtils.createGeneral(config.gramerPrefix, config.gramerSuffix);
        }
        return config;
    }
}