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
import java.util.EnumMap;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import com.github.microprograms.poi_template.exception.RenderException;
import com.github.microprograms.poi_template.policy.DocxRenderPolicy;
import com.github.microprograms.poi_template.policy.NumberingRenderPolicy;
import com.github.microprograms.poi_template.policy.PictureRenderPolicy;
import com.github.microprograms.poi_template.policy.RenderPolicy;
import com.github.microprograms.poi_template.policy.TableRenderPolicy;
import com.github.microprograms.poi_template.policy.TextRenderPolicy;
import com.github.microprograms.poi_template.policy.reference.DefaultChartTemplateRenderPolicy;
import com.github.microprograms.poi_template.policy.reference.DefaultPictureTemplateRenderPolicy;
import com.github.microprograms.poi_template.policy.reference.MultiSeriesChartTemplateRenderPolicy;
import com.github.microprograms.poi_template.policy.reference.SingleSeriesChartTemplateRenderPolicy;
import com.github.microprograms.poi_template.render.RenderContext;
import com.github.microprograms.poi_template.render.compute.DefaultRenderDataComputeFactory;
import com.github.microprograms.poi_template.render.compute.RenderDataComputeFactory;
import com.github.microprograms.poi_template.resolver.DefaultElementTemplateFactory;
import com.github.microprograms.poi_template.resolver.ElementTemplateFactory;
import com.github.microprograms.poi_template.template.ChartTemplate;
import com.github.microprograms.poi_template.template.MetaTemplate;
import com.github.microprograms.poi_template.template.PictureTemplate;
import com.github.microprograms.poi_template.util.RegexUtils;

import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;

/**
 * 模板配置
 */
public class Configure implements Cloneable {

    /**
     * regular expression: Chinese, letters, numbers and underscores
     */
    public static final String DEFAULT_GRAMER_REGEX = "((#)?[\\w\\u4e00-\\u9fa5]+(\\.[\\w\\u4e00-\\u9fa5]+)*)?";

    private CustomPolicyFinder CUSTOM_POLICY_FINDER = (tagName, configure) -> configure.getCustomPolicys().get(tagName);

    /**
     * template by bind: Highest priority
     */
    private final Map<String, RenderPolicy> CUSTOM_POLICYS = new HashMap<String, RenderPolicy>();

    /**
     * template by xwpfRun: Medium priority
     */
    private final Map<Character, RenderPolicy> DEFAULT_POLICYS = new HashMap<Character, RenderPolicy>();

    /**
     * template by xwpfchart: Medium priority
     */
    private final Map<ChartTypes, RenderPolicy> DEFAULT_CHART_POLICYS = new EnumMap<ChartTypes, RenderPolicy>(
            ChartTypes.class);

    /**
     * template by element template: Lowest priority
     */
    private final Map<Class<? extends MetaTemplate>, RenderPolicy> DEFAULT_TEMPLATE_POLICYS = new HashMap<>();

    /**
     * if & for each
     * <p>
     * eg. {{?user}} Hello, World {{/user}}
     * </p>
     */
    private Pair<Character, Character> iterable = Pair.of(GramerSymbol.ITERABLE_START.getSymbol(),
            GramerSymbol.BLOCK_END.getSymbol());

    /**
     * tag prefix
     */
    private String gramerPrefix = "{{";

    /**
     * tag suffix
     */
    private String gramerSuffix = "}}";

    /**
     * tag regular expression
     */
    private String grammerRegex = DEFAULT_GRAMER_REGEX;

    /**
     * the factory of render data compute
     */
    private RenderDataComputeFactory renderDataComputeFactory = new DefaultRenderDataComputeFactory(false);

    /**
     * the factory of resovler run template
     */
    private ElementTemplateFactory elementTemplateFactory = new DefaultElementTemplateFactory();

    /**
     * the policy of process tag for valid render data error(null or illegal)
     */
    private ValidErrorHandler handler = new ClearHandler();

    /**
     * sp el custom static method
     */
    private Map<String, Method> spELFunction = new HashMap<String, Method>();

    public Configure() {
        plugin(GramerSymbol.TEXT, new TextRenderPolicy());
        plugin(GramerSymbol.TEXT_ALIAS, new TextRenderPolicy());
        plugin(GramerSymbol.IMAGE, new PictureRenderPolicy());
        plugin(GramerSymbol.TABLE, new TableRenderPolicy());
        plugin(GramerSymbol.NUMBERING, new NumberingRenderPolicy());
        plugin(GramerSymbol.DOCX_TEMPLATE, new DocxRenderPolicy());

        RenderPolicy multiSeriesRenderPolicy = new MultiSeriesChartTemplateRenderPolicy();
        plugin(ChartTypes.AREA, multiSeriesRenderPolicy);
        plugin(ChartTypes.AREA3D, multiSeriesRenderPolicy);
        plugin(ChartTypes.BAR, multiSeriesRenderPolicy);
        plugin(ChartTypes.BAR3D, multiSeriesRenderPolicy);
        plugin(ChartTypes.LINE, multiSeriesRenderPolicy);
        plugin(ChartTypes.LINE3D, multiSeriesRenderPolicy);
        plugin(ChartTypes.RADAR, multiSeriesRenderPolicy);

        RenderPolicy singleSeriesRenderPolicy = new SingleSeriesChartTemplateRenderPolicy();
        plugin(ChartTypes.PIE, singleSeriesRenderPolicy);
        plugin(ChartTypes.PIE3D, singleSeriesRenderPolicy);
        plugin(ChartTypes.DOUGHNUT, singleSeriesRenderPolicy);

        plugin(PictureTemplate.class, new DefaultPictureTemplateRenderPolicy());
        plugin(ChartTemplate.class, new DefaultChartTemplateRenderPolicy());
    }

    public static Configure createDefault() {
        return builder().build();
    }

    public static ConfigureBuilder builder() {
        return new ConfigureBuilder();
    }

    public Configure plugin(char c, RenderPolicy policy) {
        DEFAULT_POLICYS.put(Character.valueOf(c), policy);
        return this;
    }

    public Configure plugin(GramerSymbol symbol, RenderPolicy policy) {
        DEFAULT_POLICYS.put(symbol.getSymbol(), policy);
        return this;
    }

    public Configure plugin(Class<? extends MetaTemplate> clazz, RenderPolicy policy) {
        DEFAULT_TEMPLATE_POLICYS.put(clazz, policy);
        return this;
    }

    public Configure plugin(ChartTypes chartType, RenderPolicy policy) {
        DEFAULT_CHART_POLICYS.put(chartType, policy);
        return this;
    }

    public void customPolicy(String tagName, RenderPolicy policy) {
        CUSTOM_POLICYS.put(tagName, policy);
    }

    public RenderPolicy getTemplatePolicy(Class<?> clazz) {
        return DEFAULT_TEMPLATE_POLICYS.get(clazz);
    }

    public RenderPolicy getCustomPolicy(String tagName) {
        return CUSTOM_POLICY_FINDER.getCustomPolicy(tagName, this);
    }

    public RenderPolicy getDefaultPolicy(Character sign) {
        return DEFAULT_POLICYS.get(sign);
    }

    public RenderPolicy getChartPolicy(ChartTypes type) {
        return DEFAULT_CHART_POLICYS.get(type);
    }

    public CustomPolicyFinder getCustomPolicyFinder() {
        return CUSTOM_POLICY_FINDER;
    }

    public void setCustomPolicyFinder(CustomPolicyFinder customPolicyFinder) {
        CUSTOM_POLICY_FINDER = customPolicyFinder;
    }

    public Map<Character, RenderPolicy> getDefaultPolicys() {
        return DEFAULT_POLICYS;
    }

    public Map<String, RenderPolicy> getCustomPolicys() {
        return CUSTOM_POLICYS;
    }

    public Map<ChartTypes, RenderPolicy> getChartPolicys() {
        return DEFAULT_CHART_POLICYS;
    }

    public Set<Character> getGramerChars() {
        Set<Character> ret = new HashSet<Character>(DEFAULT_POLICYS.keySet());
        // ? /
        ret.add(iterable.getKey());
        ret.add(iterable.getValue());
        return ret;
    }

    public String getGramerPrefix() {
        return gramerPrefix;
    }

    public void setGramerPrefix(String gramerPrefix) {
        this.gramerPrefix = gramerPrefix;
    }

    public String getGramerSuffix() {
        return gramerSuffix;
    }

    public void setGramerSuffix(String gramerSuffix) {
        this.gramerSuffix = gramerSuffix;
    }

    public String getGrammerRegex() {
        return grammerRegex;
    }

    public void setGrammerRegex(String grammerRegex) {
        this.grammerRegex = grammerRegex;
    }

    public ValidErrorHandler getValidErrorHandler() {
        return handler;
    }

    public void setValidErrorHandler(ValidErrorHandler handler) {
        this.handler = handler;
    }

    public RenderDataComputeFactory getRenderDataComputeFactory() {
        return renderDataComputeFactory;
    }

    public void setRenderDataComputeFactory(RenderDataComputeFactory renderDataComputeFactory) {
        this.renderDataComputeFactory = renderDataComputeFactory;
    }

    public ElementTemplateFactory getElementTemplateFactory() {
        return elementTemplateFactory;
    }

    public void setElementTemplateFactory(ElementTemplateFactory elementTemplateFactory) {
        this.elementTemplateFactory = elementTemplateFactory;
    }

    public Pair<Character, Character> getIterable() {
        return iterable;
    }

    public void setIterable(Pair<Character, Character> iterable) {
        this.iterable = iterable;
    }

    public Map<String, Method> getSpELFunction() {
        return spELFunction;
    }

    public void setSpELFunction(Map<String, Method> spELFunction) {
        this.spELFunction = spELFunction;
    }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append("Configure Info").append(":\n");
        sb.append("  Basic gramer: ").append(gramerPrefix).append(gramerSuffix).append("\n");
        sb.append("  If and foreach gramer: ").append(gramerPrefix).append(iterable.getLeft()).append(gramerSuffix);
        sb.append(gramerPrefix).append(iterable.getRight()).append(gramerSuffix).append("\n");
        sb.append("  Regex:").append(grammerRegex).append("\n");
        sb.append("  Valid Error Handler: ").append(handler.getClass().getSimpleName()).append("\n");
        sb.append("  Default Plugin: ").append("\n");
        DEFAULT_POLICYS.forEach((chara, policy) -> {
            sb.append("    ").append(gramerPrefix).append(chara.charValue()).append(gramerSuffix);
            sb.append("->").append(policy.getClass().getSimpleName()).append("\n");
        });
        sb.append("  Bind Plugin: ").append("\n");
        CUSTOM_POLICYS.forEach((str, policy) -> {
            sb.append("    ").append(gramerPrefix).append(str).append(gramerSuffix);
            sb.append("->").append(policy.getClass().getSimpleName()).append("\n");
        });
        sb.append("  Chart Plugin: ").append("\n");
        DEFAULT_CHART_POLICYS.forEach((type, policy) -> {
            sb.append("    ").append(type);
            sb.append("->").append(policy.getClass().getSimpleName()).append("\n");
        });
        sb.append("  Template Plugin: ").append("\n");
        DEFAULT_TEMPLATE_POLICYS.forEach((clazz, policy) -> {
            sb.append("    ").append(clazz.getSimpleName());
            sb.append("->").append(policy.getClass().getSimpleName()).append("\n");
        });
        sb.append("  SpELFunction: ").append("\n");
        spELFunction.forEach((str, method) -> {
            sb.append("    ").append(str);
            sb.append("->").append(method.toString()).append("\n");
        });
        return sb.toString();
    }

    @Override
    protected Configure clone() throws CloneNotSupportedException {
        return (Configure) super.clone();
    }

    public Configure copy(String prefix, String suffix) throws CloneNotSupportedException {
        Configure clone = clone();
        clone.gramerPrefix = prefix;
        clone.gramerSuffix = suffix;
        clone.grammerRegex = RegexUtils.createGeneral(clone.gramerPrefix, clone.gramerSuffix);
        return clone;
    }

    public interface ValidErrorHandler {
        void handler(RenderContext<?> context);
    }

    public static class DiscardHandler implements ValidErrorHandler {
        @Override
        public void handler(RenderContext<?> context) {
            // no-op
        }
    }

    public static class ClearHandler implements ValidErrorHandler {
        @Override
        public void handler(RenderContext<?> context) {
            context.getRun().setText("", 0);
        }
    }

    public static class AbortHandler implements ValidErrorHandler {
        @Override
        public void handler(RenderContext<?> context) {
            throw new RenderException("Non-existent variable and a variable with illegal value for "
                    + context.getTagSource() + ", data: " + context.getData());
        }
    }

    @FunctionalInterface
    public static interface CustomPolicyFinder {
        RenderPolicy getCustomPolicy(String tagName, Configure configure);
    }
}
