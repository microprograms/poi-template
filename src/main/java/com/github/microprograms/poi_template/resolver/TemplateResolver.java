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
package com.github.microprograms.poi_template.resolver;

import java.util.ArrayList;
import java.util.Deque;
import java.util.LinkedList;
import java.util.List;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.microprograms.poi_template.config.Configure;
import com.github.microprograms.poi_template.exception.ResolverException;
import com.github.microprograms.poi_template.template.BlockTemplate;
import com.github.microprograms.poi_template.template.ChartTemplate;
import com.github.microprograms.poi_template.template.ElementTemplate;
import com.github.microprograms.poi_template.template.IterableTemplate;
import com.github.microprograms.poi_template.template.MetaTemplate;
import com.github.microprograms.poi_template.template.PictureTemplate;
import com.github.microprograms.poi_template.template.run.RunTemplate;
import com.github.microprograms.poi_template.util.ReflectionUtils;
import com.github.microprograms.poi_template.xwpf.CTDrawingWrapper;
import com.github.microprograms.poi_template.xwpf.XWPFRunWrapper;

/**
 * Resolver
 */
public class TemplateResolver extends AbstractResolver {

    private static Logger logger = LoggerFactory.getLogger(TemplateResolver.class);

    private ElementTemplateFactory elementTemplateFactory;

    public TemplateResolver(Configure config) {
        this(config, config.getElementTemplateFactory());
    }

    private TemplateResolver(Configure config, ElementTemplateFactory elementTemplateFactory) {
        super(config);
        this.elementTemplateFactory = elementTemplateFactory;
    }

    @Override
    public List<MetaTemplate> resolveDocument(XWPFDocument doc) {
        List<MetaTemplate> metaTemplates = new ArrayList<>();
        if (null == doc) return metaTemplates;
        logger.info("Resolve the document start...");
        metaTemplates.addAll(resolveBodyElements(doc.getBodyElements()));
        metaTemplates.addAll(resolveHeaders(doc.getHeaderList()));
        metaTemplates.addAll(resolveFooters(doc.getFooterList()));
        logger.info("Resolve the document end, resolve and create {} MetaTemplates.", metaTemplates.size());
        return metaTemplates;
    }

    @Override
    public List<MetaTemplate> resolveBodyElements(List<IBodyElement> bodyElements) {
        List<MetaTemplate> metaTemplates = new ArrayList<>();
        if (null == bodyElements) return metaTemplates;

        // current iterable templates state
        Deque<BlockTemplate> stack = new LinkedList<BlockTemplate>();

        for (IBodyElement element : bodyElements) {
            if (element == null) continue;
            if (element.getElementType() == BodyElementType.PARAGRAPH) {
                XWPFParagraph paragraph = (XWPFParagraph) element;
                new RunningRunParagraph(paragraph, templatePattern).refactorRun();
                resolveXWPFRuns(paragraph.getRuns(), metaTemplates, stack);
            } else if (element.getElementType() == BodyElementType.TABLE) {
                XWPFTable table = (XWPFTable) element;
                List<XWPFTableRow> rows = table.getRows();
                if (null == rows) continue;
                for (XWPFTableRow row : rows) {
                    List<XWPFTableCell> cells = row.getTableCells();
                    if (null == cells) continue;
                    cells.forEach(cell -> {
                        addNewMeta(metaTemplates, stack, resolveBodyElements(cell.getBodyElements()));
                    });
                }
            }
        }

        checkStack(stack);
        return metaTemplates;
    }

    @Override
    public List<MetaTemplate> resolveXWPFRuns(List<XWPFRun> runs) {
        List<MetaTemplate> metaTemplates = new ArrayList<>();
        if (runs == null) return metaTemplates;

        Deque<BlockTemplate> stack = new LinkedList<BlockTemplate>();
        resolveXWPFRuns(runs, metaTemplates, stack);
        checkStack(stack);
        return metaTemplates;
    }

    private void resolveXWPFRuns(List<XWPFRun> runs, final List<MetaTemplate> metaTemplates,
            final Deque<BlockTemplate> stack) {
        for (XWPFRun run : runs) {
            String text = null;
            if (StringUtils.isBlank(text = run.getText(0))) {
                // textbox
                List<MetaTemplate> visitBodyElements = resolveTextbox(run);
                if (!visitBodyElements.isEmpty()) {
                    addNewMeta(metaTemplates, stack, visitBodyElements);
                    continue;
                }

                List<PictureTemplate> pictureTemplates = resolveXWPFPictures(run.getEmbeddedPictures());
                if (!pictureTemplates.isEmpty()) {
                    addNewMeta(metaTemplates, stack, pictureTemplates);
                    continue;
                }

                ChartTemplate chartTemplate = resolveXWPFChart(run);
                if (null != chartTemplate) {
                    addNewMeta(metaTemplates, stack, chartTemplate);
                    continue;
                }
                continue;
            }
            RunTemplate runTemplate = (RunTemplate) parseTemplateFactory(text, run, run);
            if (null == runTemplate) continue;
            char charValue = runTemplate.getSign().charValue();
            if (charValue == config.getIterable().getLeft()) {
                IterableTemplate freshIterableTemplate = new IterableTemplate(runTemplate);
                stack.push(freshIterableTemplate);
            } else if (charValue == config.getIterable().getRight()) {
                if (stack.isEmpty()) throw new ResolverException(
                        "Mismatched start/end tags: No start mark found for end mark " + runTemplate);
                BlockTemplate latestIterableTemplate = stack.pop();
                if (StringUtils.isNotEmpty(runTemplate.getTagName())
                        && !latestIterableTemplate.getStartMark().getTagName().equals(runTemplate.getTagName())) {
                    throw new ResolverException("Mismatched start/end tags: start mark "
                            + latestIterableTemplate.getStartMark() + " does not match to end mark " + runTemplate);
                }
                latestIterableTemplate.setEndMark(runTemplate);
                if (latestIterableTemplate instanceof IterableTemplate) {
                    latestIterableTemplate = ((IterableTemplate) latestIterableTemplate).buildIfInline();
                }
                addNewMeta(metaTemplates, stack, latestIterableTemplate);
            } else {
                addNewMeta(metaTemplates, stack, runTemplate);
            }
        }
    }

    private ChartTemplate resolveXWPFChart(XWPFRun run) {
        CTDrawing ctDrawing = getCTDrawing(run);
        if (null == ctDrawing) return null;
        CTDrawingWrapper wrapper = new CTDrawingWrapper(ctDrawing);
        String title = wrapper.getTitle();
        if (null == title) return null;
        String rid = wrapper.getCharId();
        if (null == rid) return null;
        POIXMLDocumentPart documentPart = run.getDocument().getRelationById(rid);
        if (null == documentPart || !(documentPart instanceof XWPFChart)) return null;
        return (ChartTemplate) parseTemplateFactory(title, (XWPFChart) documentPart, run);
    }

    private List<PictureTemplate> resolveXWPFPictures(List<XWPFPicture> embeddedPictures) {
        List<PictureTemplate> metaTemplates = new ArrayList<>();
        if (embeddedPictures == null) return metaTemplates;

        for (XWPFPicture pic : embeddedPictures) {
            // it's array, to do in the future
            CTDrawing ctDrawing = getCTDrawing(pic);
            if (null == ctDrawing) continue;
            CTDrawingWrapper wrapper = new CTDrawingWrapper(ctDrawing);
            String title = wrapper.getTitle();
            if (null == title) continue;
            PictureTemplate pictureTemplate = (PictureTemplate) parseTemplateFactory(title, pic, null);
            if (null != pictureTemplate) metaTemplates.add(pictureTemplate);
        }
        return metaTemplates;
    }

    private CTDrawing getCTDrawing(XWPFPicture pic) throws RuntimeException {
        XWPFRun run = (XWPFRun) ReflectionUtils.getValue("run", pic);
        return getCTDrawing(run);
    }

    private CTDrawing getCTDrawing(XWPFRun run) {
        CTR ctr = run.getCTR();
        CTDrawing ctDrawing = CollectionUtils.isNotEmpty(ctr.getDrawingList()) ? ctr.getDrawingArray(0) : null;
        return ctDrawing;
    }

    private void addNewMeta(final List<MetaTemplate> metaTemplates, final Deque<BlockTemplate> stack,
            List<? extends MetaTemplate> newMeta) {
        if (stack.isEmpty()) {
            metaTemplates.addAll(newMeta);
        } else {
            stack.peek().getTemplates().addAll(newMeta);
        }
    }

    private <T extends MetaTemplate> void addNewMeta(final List<MetaTemplate> metaTemplates,
            final Deque<BlockTemplate> stack, T newMeta) {
        if (stack.isEmpty()) {
            metaTemplates.add(newMeta);
        } else {
            stack.peek().getTemplates().add(newMeta);
        }
    }

    private void checkStack(Deque<BlockTemplate> stack) {
        if (!stack.isEmpty()) {
            throw new ResolverException(
                    "Mismatched start/end tags: No end iterable mark found for start mark " + stack.peek());
        }
    }

    private List<MetaTemplate> resolveTextbox(XWPFRun run) {
        XWPFRunWrapper runWrapper = new XWPFRunWrapper(run);
        if (null == runWrapper.getWpstxbx()) return new ArrayList<>();
        return resolveBodyElements(runWrapper.getWpstxbx().getBodyElements());
    }

    List<MetaTemplate> resolveHeaders(List<XWPFHeader> headers) {
        List<MetaTemplate> metaTemplates = new ArrayList<>();
        if (null == headers) return metaTemplates;

        headers.forEach(header -> {
            metaTemplates.addAll(resolveBodyElements(header.getBodyElements()));
        });
        return metaTemplates;
    }

    List<MetaTemplate> resolveFooters(List<XWPFFooter> footers) {
        List<MetaTemplate> metaTemplates = new ArrayList<>();
        if (null == footers) return metaTemplates;

        footers.forEach(footer -> {
            metaTemplates.addAll(resolveBodyElements(footer.getBodyElements()));
        });
        return metaTemplates;
    }

    ElementTemplate parseTemplateFactory(String text, Object obj, XWPFRun run) {
        if (templatePattern.matcher(text).matches()) {
            logger.debug("Resolve where text: {}, and create ElementTemplate for {}", text, obj.getClass());
            String tag = gramerPattern.matcher(text).replaceAll("").trim();
            if (obj.getClass() == XWPFRun.class) {
                return (RunTemplate) elementTemplateFactory.createRunTemplate(tag, (XWPFRun) obj);
            } else if (obj.getClass() == XWPFPicture.class) {
                return (PictureTemplate) elementTemplateFactory.createPicureTemplate(tag, (XWPFPicture) obj);
            } else if (obj.getClass() == XWPFChart.class) {
                return (ChartTemplate) elementTemplateFactory.createChartTemplate(tag, (XWPFChart) obj, run);
            }
        }
        return null;
    }

}
