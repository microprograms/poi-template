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
import java.util.Map;

import com.github.microprograms.poi_template.config.Configure.CustomPolicyFinder;
import com.github.microprograms.poi_template.config.Configure.ValidErrorHandler;
import com.github.microprograms.poi_template.policy.RenderPolicy;
import com.github.microprograms.poi_template.render.compute.DefaultELRenderDataCompute;
import com.github.microprograms.poi_template.render.compute.RenderDataComputeFactory;
import com.github.microprograms.poi_template.render.compute.SpELRenderDataCompute;
import com.github.microprograms.poi_template.resolver.ElementTemplateFactory;
import com.github.microprograms.poi_template.template.MetaTemplate;
import com.github.microprograms.poi_template.util.RegexUtils;

import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;

public class ConfigureBuilder {
    private Configure config = new Configure();
    private boolean usedSpringEL;
    private boolean changeRegex;

    public ConfigureBuilder gramer(String prefix, String suffix) {
        config.setGramerPrefix(prefix);
        config.setGramerSuffix(suffix);
        return this;
    }

    public ConfigureBuilder iterableLeft(char c) {
        config.setIterable(Pair.of(c, config.getIterable().getRight()));
        return this;
    }

    public ConfigureBuilder grammerRegex(String reg) {
        changeRegex = true;
        config.setGrammerRegex(reg);
        return this;
    }

    public ConfigureBuilder useSpringEL() {
        return useSpringEL(root -> new SpELRenderDataCompute(root));
    }

    public ConfigureBuilder useSpringEL(Map<String, Method> spELFunction) {
        return useSpringEL(root -> new SpELRenderDataCompute(root, spELFunction));
    }

    public ConfigureBuilder useSpringEL(RenderDataComputeFactory renderDataComputeFactory) {
        usedSpringEL = true;
        return _setRenderDataComputeFactory(renderDataComputeFactory);
    }

    public ConfigureBuilder useDefaultStrictEL() {
        return _setRenderDataComputeFactory(root -> new DefaultELRenderDataCompute(root, true));
    }

    public ConfigureBuilder useDefaultEL() {
        return _setRenderDataComputeFactory(root -> new DefaultELRenderDataCompute(root, false));
    }

    public ConfigureBuilder validErrorHandler(ValidErrorHandler handler) {
        config.setValidErrorHandler(handler);
        return this;
    }

    private ConfigureBuilder _setRenderDataComputeFactory(RenderDataComputeFactory renderDataComputeFactory) {
        config.setRenderDataComputeFactory(renderDataComputeFactory);
        return this;
    }

    public ConfigureBuilder elementTemplateFactory(ElementTemplateFactory elementTemplateFactory) {
        config.setElementTemplateFactory(elementTemplateFactory);
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

    public ConfigureBuilder bind(String tagName, RenderPolicy policy) {
        config.customPolicy(tagName, policy);
        return this;
    }

    public ConfigureBuilder customPolicyFinder(CustomPolicyFinder customPolicyFinder) {
        config.setCustomPolicyFinder(customPolicyFinder);
        return this;
    }

    public Configure build() {
        if (usedSpringEL && !changeRegex) {
            config.setGramerPrefix(RegexUtils.createGeneral(config.getGramerPrefix(), config.getGramerSuffix()));
        }
        return config;
    }
}