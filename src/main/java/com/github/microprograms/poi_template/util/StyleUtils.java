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
package com.github.microprograms.poi_template.util;

import java.math.BigInteger;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.LineSpacingRule;
import org.apache.poi.xwpf.usermodel.TableWidthType;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHeight;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHighlight;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHpsMeasure;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTInd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTParaRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGridCol;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTUnderline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHeightRule;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHighlightColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHighlightColor.Enum;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STUnderline;

import com.github.microprograms.poi_template.data.style.CellStyle;
import com.github.microprograms.poi_template.data.style.ParagraphStyle;
import com.github.microprograms.poi_template.data.style.RowStyle;
import com.github.microprograms.poi_template.data.style.Style;
import com.github.microprograms.poi_template.data.style.Style.StyleBuilder;
import com.github.microprograms.poi_template.data.style.TableStyle;

/**
 * set style for run, paragraph, table...
 */
public final class StyleUtils {

    /**
     * set run style by style
     * 
     * @param run
     * @param style
     */
    public static void styleRun(XWPFRun run, Style style) {
        if (null == run || null == style) return;
        String color = style.getColor();
        CTRPr pr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
        if (StringUtils.isNotBlank(color)) {
            // run.setColor(color);
            // issue 326
            CTColor ctColor = pr.isSetColor() ? pr.getColor() : pr.addNewColor();
            ctColor.setVal(color);
            if (ctColor.isSetThemeColor()) ctColor.unsetThemeColor();
        }
        int fontSize = style.getFontSize();
        if (0 != fontSize && -1 != fontSize) run.setFontSize(fontSize);
        String fontFamily = style.getFontFamily();
        if (StringUtils.isNotBlank(fontFamily)) {
            run.setFontFamily(fontFamily);
            CTFonts fonts = pr.isSetRFonts() ? pr.getRFonts() : pr.addNewRFonts();
            fonts.setAscii(fontFamily);
            fonts.setHAnsi(fontFamily);
            fonts.setCs(fontFamily);
            fonts.setEastAsia(fontFamily);
        }
        Enum highlightColor = style.getHighlightColor();
        if (null != highlightColor) {
            CTHighlight highlight = pr.isSetHighlight() ? pr.getHighlight() : pr.addNewHighlight();
            STHighlightColor hColor = highlight.xgetVal();
            if (hColor == null) {
                hColor = STHighlightColor.Factory.newInstance();
            }
            STHighlightColor.Enum val = STHighlightColor.Enum.forString(highlightColor.toString());
            if (val != null) {
                hColor.setStringValue(val.toString());
                highlight.xsetVal(hColor);
            }
        }
        Boolean bold = style.isBold();
        if (null != bold) run.setBold(bold);
        Boolean italic = style.isItalic();
        if (null != italic) run.setItalic(italic);
        Boolean strike = style.isStrike();
        if (null != strike) run.setStrikeThrough(strike);
        Boolean underLine = style.isUnderLine();
        if (Boolean.TRUE.equals(underLine)) {
            run.setUnderline(UnderlinePatterns.SINGLE);
        }
        int twips = style.getCharacterSpacing();
        // in twentieths of a point
        if (0 != twips && -1 != twips) run.setCharacterSpacing(20 * twips);
        String vertAlign = style.getVertAlign();
        if (StringUtils.isNotBlank(vertAlign)) {
            run.setVerticalAlignment(vertAlign);
        }
    }

    /**
     * set run style by other run
     * 
     * @param dest
     * @param src
     */
    public static void styleRun(XWPFRun dest, XWPFRun src) {
        if (null == dest || null == src) return;
        if (StringUtils.isNotEmpty(src.getStyle())) dest.setStyle(src.getStyle());
        if (Boolean.TRUE.equals(src.isBold())) dest.setBold(src.isBold());
        if (StringUtils.isNotBlank(src.getColor())) dest.setColor(src.getColor());
        if (0 != src.getCharacterSpacing()) dest.setCharacterSpacing(src.getCharacterSpacing());
        if (StringUtils.isNotBlank(src.getFontFamily())) dest.setFontFamily(src.getFontFamily());
        int fontSize = src.getFontSize();
        if (-1 != fontSize) dest.setFontSize(fontSize);
        if (Boolean.TRUE.equals(src.isItalic())) dest.setItalic(src.isItalic());
        if (Boolean.TRUE.equals(src.isStrikeThrough())) dest.setStrikeThrough(src.isStrikeThrough());
        if (UnderlinePatterns.NONE != src.getUnderline()) dest.setUnderline(src.getUnderline());
    }

    /**
     * set paragraph style
     * 
     * @param paragraph
     * @param style
     */
    public static void styleParagraph(XWPFParagraph paragraph, ParagraphStyle style) {
        if (null == paragraph || null == style) return;
        stylePpr(paragraph, style);
        styleParaRpr(paragraph, style.getGlyphStyle());
    }

    /**
     * set table style
     * 
     * @param table
     * @param tableStyle
     */
    public static void styleTable(XWPFTable table, TableStyle tableStyle) {
        if (null == table || null == tableStyle) return;
        String width = tableStyle.getWidth();
        ensureTblW(table);
        table.setWidth(width);

        int[] colWidths = tableStyle.getColWidths();
        if (null == colWidths && table.getWidthType() == TableWidthType.DXA) {
            colWidths = UnitUtils.average(Integer.valueOf(width), TableTools.obtainColumnSize(table));
            // TODO support calc col width of pct and auto for Apple Pages!
        }
        if (null != colWidths) {
            CTTblGrid tblGrid = TableTools.getTblGrid(table);
            CTTblLayoutType tblLayout = TableTools.getTblLayout(table);
            tblLayout.setType(STTblLayoutType.FIXED);
            for (int index = 0; index < colWidths.length; index++) {
                CTTblGridCol addNewGridCol = tblGrid.addNewGridCol();
                addNewGridCol.setW(BigInteger.valueOf(colWidths[index]));

                List<XWPFTableRow> rows = table.getRows();
                for (XWPFTableRow row : rows) {
                    row.getCell(index).setWidth(colWidths[index] + "");
                }
            }
        }

        TableTools.setBorder(table::setLeftBorder, tableStyle.getLeftBorder());
        TableTools.setBorder(table::setRightBorder, tableStyle.getRightBorder());
        TableTools.setBorder(table::setTopBorder, tableStyle.getTopBorder());
        TableTools.setBorder(table::setBottomBorder, tableStyle.getBottomBorder());
        TableTools.setBorder(table::setInsideHBorder, tableStyle.getInsideHBorder());
        TableTools.setBorder(table::setInsideVBorder, tableStyle.getInsideVBorder());

        if (null != tableStyle.getAlign()) {
            table.setTableAlignment(tableStyle.getAlign());
        }

    }

    private static void ensureTblW(XWPFTable table) {
        CTTbl ctTbl = table.getCTTbl();
        CTTblPr tblPr = (ctTbl.getTblPr() != null) ? ctTbl.getTblPr() : ctTbl.addNewTblPr();
        if (!tblPr.isSetTblW()) tblPr.addNewTblW();
    }

    /**
     * set row style
     * 
     * @param row
     * @param rowStyle
     */
    public static void styleTableRow(XWPFTableRow row, RowStyle rowStyle) {
        if (null == row || null == rowStyle) return;
        int height = rowStyle.getHeight();
        if (0 != height) {
            row.setHeight(height);
            CTRow ctRow = row.getCtRow();
            CTTrPr properties = (ctRow.isSetTrPr()) ? ctRow.getTrPr() : ctRow.addNewTrPr();
            CTHeight h = properties.sizeOfTrHeightArray() == 0 ? properties.addNewTrHeight()
                    : properties.getTrHeightArray(0);
            h.setHRule(STHeightRule.AT_LEAST);
        }
    }

    /**
     * set cell style
     * 
     * @param cell
     * @param cellStyle
     */
    public static void styleTableCell(XWPFTableCell cell, CellStyle cellStyle) {
        if (null == cell || null == cellStyle) return;
        if (null != cellStyle.getVertAlign()) {
            cell.setVerticalAlignment(cellStyle.getVertAlign());
        }
        if (null != cellStyle.getBackgroundColor()) {
            cell.setColor(cellStyle.getBackgroundColor());
        }
    }

    /**
     * set w:rPr style
     * 
     * @param pr
     * @param style
     */
    private static void styleParaRpr(CTParaRPr pr, Style style) {
        if (null == pr || null == style) return;
        if (StringUtils.isNotBlank(style.getColor())) {
            CTColor color = pr.isSetColor() ? pr.getColor() : pr.addNewColor();
            color.setVal(style.getColor());
        }

        if (null != style.isItalic()) {
            CTOnOff italic = pr.isSetI() ? pr.getI() : pr.addNewI();
            italic.setVal(style.isItalic() ? STOnOff.TRUE : STOnOff.FALSE);
        }

        if (null != style.isBold()) {
            CTOnOff bold = pr.isSetB() ? pr.getB() : pr.addNewB();
            bold.setVal(style.isBold() ? STOnOff.TRUE : STOnOff.FALSE);
        }

        if (0 != style.getFontSize() && -1 != style.getFontSize()) {
            BigInteger bint = new BigInteger("" + style.getFontSize());
            CTHpsMeasure ctSize = pr.isSetSz() ? pr.getSz() : pr.addNewSz();
            ctSize.setVal(bint.multiply(new BigInteger("2")));
        }

        if (null != style.isStrike()) {
            CTOnOff strike = pr.isSetStrike() ? pr.getStrike() : pr.addNewStrike();
            strike.setVal(style.isStrike() ? STOnOff.TRUE : STOnOff.FALSE);
        }

        if (Boolean.TRUE.equals(style.isUnderLine())) {
            CTUnderline underline = pr.isSetU() ? pr.getU() : pr.addNewU();
            underline.setVal(STUnderline.SINGLE);
        }

        if (StringUtils.isNotBlank(style.getFontFamily())) {
            CTFonts fonts = pr.isSetRFonts() ? pr.getRFonts() : pr.addNewRFonts();
            String fontFamily = style.getFontFamily();
            fonts.setAscii(fontFamily);
            if (!fonts.isSetHAnsi()) {
                fonts.setHAnsi(fontFamily);
            }
            if (!fonts.isSetCs()) {
                fonts.setCs(fontFamily);
            }
            if (!fonts.isSetEastAsia()) {
                fonts.setEastAsia(fontFamily);
            }
        }
    }

    public static void styleParaRpr(XWPFParagraph paragraph, Style style) {
        if (null == paragraph || null == style) return;
        CTP ctp = paragraph.getCTP();
        CTPPr pPr = ctp.isSetPPr() ? ctp.getPPr() : ctp.addNewPPr();
        CTParaRPr pr = pPr.isSetRPr() ? pPr.getRPr() : pPr.addNewRPr();
        StyleUtils.styleParaRpr(pr, style);
    }

    public static void stylePpr(XWPFParagraph paragraph, ParagraphStyle style) {
        if (null == paragraph || null == style) return;
        if (null != style.getAlign()) {
            paragraph.setAlignment(style.getAlign());
        }

        if (0 != style.getSpacing()) {
            paragraph.setSpacingBetween(style.getSpacing(), LineSpacingRule.AUTO);
        }

        CTP ctp = paragraph.getCTP();
        CTPPr pr = ctp.isSetPPr() ? ctp.getPPr() : ctp.addNewPPr();
        CTInd indent = pr.isSetInd() ? pr.getInd() : pr.addNewInd();
        if (0 != style.getIndentLeftChars()) {
            BigInteger bi = new BigInteger(String.valueOf(Math.round(style.getIndentLeftChars() * 100.0)));
            indent.setLeftChars(bi);
            if (indent.isSetLeft()) indent.unsetLeft();
        }
        if (0 != style.getIndentRightChars()) {
            BigInteger bi = new BigInteger(String.valueOf(Math.round(style.getIndentRightChars() * 100.0)));
            indent.setRightChars(bi);
            if (indent.isSetRight()) indent.unsetRight();
        }
        if (0 != style.getIndentHangingChars()) {
            BigInteger bi = new BigInteger(String.valueOf(Math.round(style.getIndentHangingChars() * 100.0)));
            indent.setHangingChars(bi);
            if (indent.isSetHanging()) indent.unsetHanging();
        }
        if (-1 != style.getNumId()) {
            paragraph.setNumID(BigInteger.valueOf(style.getNumId()));
        }
        if (-1 != style.getLvl()) {
            paragraph.setNumILvl(BigInteger.valueOf(style.getLvl()));
        }
    }

    public static Style retriveStyle(XWPFRun run) {
        if (null == run) return null;
        StyleBuilder builder = Style.builder().buildColor(run.getColor()).buildFontFamily(run.getFontFamily())
                .buildFontSize(run.getFontSize());
        if (run.isBold()) builder.buildBold();
        if (run.isItalic()) builder.buildItalic();
        if (run.isStrikeThrough()) builder.buildStrike();
        return builder.build();
    }

}
