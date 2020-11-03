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

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map.Entry;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.github.microprograms.poi_template.data.CellRenderData;
import com.github.microprograms.poi_template.data.MergeCellRule;
import com.github.microprograms.poi_template.data.MergeCellRule.Grid;
import com.github.microprograms.poi_template.data.ParagraphRenderData;
import com.github.microprograms.poi_template.data.RowRenderData;
import com.github.microprograms.poi_template.data.TableRenderData;
import com.github.microprograms.poi_template.data.style.CellStyle;
import com.github.microprograms.poi_template.data.style.ParagraphStyle;
import com.github.microprograms.poi_template.data.style.TableStyle;
import com.github.microprograms.poi_template.render.RenderContext;
import com.github.microprograms.poi_template.util.StyleUtils;
import com.github.microprograms.poi_template.util.TableTools;
import com.github.microprograms.poi_template.xwpf.BodyContainer;
import com.github.microprograms.poi_template.xwpf.BodyContainerFactory;

/**
 * table render
 */
public class TableRenderPolicy extends AbstractRenderPolicy<TableRenderData> {

    @Override
    protected boolean validate(TableRenderData data) {
        if (null == data || 0 == data.obtainColSize()) return false;
        List<RowRenderData> rows = data.getRows();
        final int col = data.obtainColSize();
        for (RowRenderData row : rows) {
            if (col != row.obtainColSize()) {
                throw new IllegalArgumentException("Number of cells in each row should be the same!");
            }
        }
        TableStyle tableStyle = data.getTableStyle();
        if (null != tableStyle) {
            int[] colWidths = tableStyle.getColWidths();
            if (null != colWidths && colWidths.length != col) {
                throw new IllegalArgumentException(
                        "The length of Colwidth array and number of columns must be the same!");
            }
        }
        return true;
    }

    @Override
    public void doRender(RenderContext<TableRenderData> context) throws Exception {
        Helper.renderTable(context.getRun(), context.getData());
    }

    @Override
    protected void afterRender(RenderContext<TableRenderData> context) {
        clearPlaceholder(context, true);
    }

    public static class Helper {

        public static void renderTable(XWPFRun run, TableRenderData data) throws Exception {
            BodyContainer bodyContainer = BodyContainerFactory.getBodyContainer(run);
            XWPFTable table = bodyContainer.insertNewTable(run, data.obtainRowSize(), data.obtainColSize());
            StyleUtils.styleTable(table, data.getTableStyle());

            int size = table.getRows().size();
            for (int i = 0; i < size; i++) {
                renderRow(table.getRows().get(i), data.getRows().get(i));
            }

            applyMergeRule(table, data.getMergeRule());
        }

        public static void renderRow(XWPFTableRow row, RowRenderData data) throws Exception {
            if (null == data) return;
            int size = row.getTableCells().size();
            if (size != data.obtainColSize()) {
                throw new IllegalArgumentException("Number of cells and render data should be the same!");
            }
            StyleUtils.styleTableRow(row, data.getRowStyle());

            CellStyle defaultCellStyle = null == data.getRowStyle() ? null : data.getRowStyle().getDefaultCellStyle();
            for (int i = 0; i < size; i++) {
                renderCell(row.getCell(i), data.getCells().get(i), defaultCellStyle);
            }
        }

        public static void renderCell(XWPFTableCell cell, CellRenderData data, CellStyle defaultCellStyle)
                throws Exception {
            if (null == data) return;
            StyleUtils.styleTableCell(cell, defaultCellStyle);
            StyleUtils.styleTableCell(cell, data.getCellStyle());

            List<ParagraphStyle> defaultParaStyles = new ArrayList<>();
            if (null != defaultCellStyle) {
                defaultParaStyles.add(defaultCellStyle.getDefaultParagraphStyle());
            } else if (null != data.getCellStyle()) {
                defaultParaStyles.add(data.getCellStyle().getDefaultParagraphStyle());
            }

            List<ParagraphRenderData> contents = data.getParagraphs();
            if (null != contents && !contents.isEmpty()) {
                BodyContainer bodyContainer = BodyContainerFactory.getBodyContainer(cell);
                XWPFParagraph placeHolder = cell.getParagraphArray(0);
                if (null == placeHolder) placeHolder = cell.addParagraph();

                for (ParagraphRenderData item : contents) {
                    XWPFParagraph paragraph = bodyContainer.insertNewParagraph(placeHolder.getCTP().newCursor());
                    ParagraphRenderPolicy.Helper.renderParagraph(paragraph.createRun(), item, defaultParaStyles);
                }

                List<XWPFParagraph> paragraphs = cell.getParagraphs();
                int pos = paragraphs.indexOf(placeHolder);
                if (-1 != pos) cell.removeParagraph(pos);
            }
        }

        private static void applyMergeRule(XWPFTable table, MergeCellRule mergeRule) {
            if (null == mergeRule) return;
            byte[][] markRemovedCell = new byte[TableTools.obtainRowSize(table)][TableTools.obtainColumnSize(table)];
            Iterator<Entry<Grid, Grid>> iterator = mergeRule.mappingIterator();
            while (iterator.hasNext()) {
                Entry<Grid, Grid> next = iterator.next();
                Grid key = next.getKey();
                Grid value = next.getValue();
                int startI = key.getI() > value.getI() ? value.getI() : key.getI();
                int startJ = key.getJ() > value.getJ() ? value.getJ() : key.getJ();
                int endI = key.getI() > value.getI() ? key.getI() : value.getI();
                int endJ = key.getJ() > value.getJ() ? key.getJ() : value.getJ();
                // merge(VMerge mark) vertical
                if (startI != endI) {
                    for (int j = startJ; j <= endJ; j++) {
                        TableTools.mergeCellsVertically(table, j, startI, endI);
                    }
                }
                // merge horizontal cells without remove cells
                if (startJ != endJ) {
                    for (int i = startI; i <= endI; i++) {
                        TableTools.mergeCellsHorizontalWithoutRemove(table, i, startJ, endJ);
                        for (int removedCol = startJ + 1; removedCol <= endJ; removedCol++) {
                            markRemovedCell[i][removedCol] = 1;
                        }
                    }
                }
            }
            // remove marked cell safely
            for (int i = 0; i < markRemovedCell.length; i++) {
                for (int j = markRemovedCell[i].length - 1; j >= 0; j--) {
                    if (markRemovedCell[i][j] >= 1) {
                        table.getRow(i).removeCell(j);
                        if (table.getRow(i).getTableCells().size() != table.getRow(i).getCtRow().sizeOfTcArray()) {
                            table.getRow(i).getCtRow().removeTc(j);
                        }
                    }
                }
            }
        }

    }

}